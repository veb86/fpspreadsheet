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
  Writes the conditional format "color range": The cell background in the
  specified cell range is painted in colors according to the cell values.

  The colors are varied between the given start color (at the miniumum cell value)
  and end color (at the maximum cell value).

  @param ARange      Range of the cells to be formatted
  @param AStartColor Start color of the gradient of cell background colors corresponding to the minimum cell value in the given cell range.
  @param AEndColor   End color of the gradient of cell background colors corresponding to the maximum cell value in the given cell range.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor, AEndColor: TsColor): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, AEndColor);
  StoreCFIndexInCells(Self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "color range": The cell background in the
  specified cell range is painted in colors according to the cell values.

  The 3-color gradient varies between the given start color (at the minimum cell value),
  the center color (at the median value) and the end color (at the maximum cell value).

  @param  ARange       Range of the cells to be formatted
  @param  AStartColor  Start color of the background color gradient corresponding to the minimum cell value in the given cell range.
  @param  ACenterColor Intermediate color, corresponding to the median of the cell values.
  @param  AEndColor    End color, corresponding to the maximum cell value.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor, ACenterColor, AEndColor: TsColor): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, ACenterColor, AEndColor);
  StoreCFIndexInCells(Self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "color range": The cell background in the
  specified cell range is painted in colors according to the cell values.

  The 2-color background color gradient begins at the cell value AStartValue with
  the AStartColor and is valid for the cells fulfilling the condition AStartKind.
  The gradient ends at the cell value AEndValue with AEndColor for the cells
  fulfilling the condition AEndKind.

  @param  ARange       Range of the cells to be formatted
  @param  AStartColor  First color of the background color gradient
  @param  AStartKind   Condition for the cell(s) to be painted in AStartColor
  @param  AStartValue  Value to be used in the AStartKind condition
  @param  AEndColor    Last color of the background color gradient
  @param  AEndKind     Conditions for the cell(s) to be painted in AEndColor
  @param  AEndValue    Value to be used in the AEndKind condition

  @seeAlso TsCFValueKind
-------------------------------------------------------------------------------}
function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
  AEndColor: TsColor; AEndKind: TsCFValueKind; AEndValue: Double): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, AStartKind, AStartValue,
    AEndColor, AEndKind, AEndValue);
  StoreCFIndexInCells(Self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "color range"

  Writes the conditional format "color range": The cell background in the
  specified cell range is painted in colors according to the cell values.

  The 3-color background color gradient begins at the cell value AStartValue with
  the AStartColor and is valid for the cells fulfilling the condition AStartKind.
  The ACenterColor is an intermediate color for the cell(s) fulfilling the
  condition ACenterKind when their value is equal to ACenterValue.
  The gradient ends at the cell value AEndValue with AEndColor for the cells
  fulfilling the condition AEndKind.

  @param  ARange       Range of the cells to be formatted
  @param  AStartColor  First color of the background color gradient
  @param  AStartKind   Condition for the cell(s) to be painted in AStartColor
  @param  AStartValue  Value to be used in the AStartKind condition
  @param  ACenterColor Intermediate color of the background color gradient
  @param  ACenterKind  Condition for the cell(s) to be painted in ACenterColor
  @param  ACenterValue Value to be used in the ACenterKind condition
  @param  AEndColor    Last color of the background color gradient
  @param  AEndKind     Conditions for the cell(s) to be painted in AEndColor
  @param  AEndValue    Value to be used in the AEndKind condition

  @seeAlso TsCFValueKind
-------------------------------------------------------------------------------}
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
  Writes the conditional format "data bars" in which the cell values are
  illustrated graphically as horizontal bars.

  @param  ARange     Range of the cells to be formatted
  @param  ABarColor  Color of the bars
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDataBars(ARange: TsCellRange; ABarColor: TsColor): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddDataBarRule(Self, ARange, ABarColor);
  StoreCFIndexInCells(self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "data bars" in which the cell values are
  illustrated graphically as horizontal bars.

  @param  ARange      Range of the cells to be formatted
  @param  ABarColor   Color of the bars
  @param  AStartKind  Condition to be met for the cell(s) at which the bar starts
  @param  AStartValue Value to be used for comparison in AStartKind
  @param  AEndKind    Condition to be met for the cell(s) at which the bar ends
  @param  AEndValue   Value to be used for comparison in AEndKind
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDataBars(ARange: TsCellRange; ABarColor: TsColor;
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
  Writes the conditional format "icon set" in which cell values are illustrated
  by means of icons.

  @param  ARange     Range of the cells to be formatted
  @param  AIconSet   Selection of the icon set to be used.
  @param  AHideText  Hides the normal cell text
  @param  AReverse   Icons are assigned in reverse order.

  @seeAlso  TsCFIconSet
-------------------------------------------------------------------------------}
function TsWorksheet.WriteIconSet(ARange: TsCellRange; AIconSet: TsCFIconSet;
  AHideText: Boolean = false; AReverse: Boolean = false): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddIconSetRule(Self, ARange,
    AIconSet,
    AHideText, AReverse);
  StoreCFIndexInCells(self, Result, ARange);
end;

{@@ ----------------------------------------------------------------------------
  Writes the conditional format "icon set" in which cell values are illustrated
  by means of three icons.

  @param  ARange       Range of the cells to be formatted
  @param  AIconSet     Selection of the icon set to be used.
  @param  AValueKind1  Condition for the cell values to be displayed with the first icon in the icon set
  @param  AValue1      Value to be compared with in the AValueKind1 condition
  @param  AValueKind2  Condition for the cell values to be displayed with the second icon in the icon set
  @param  AValue2      Value to be compared with in the AValueKind2 condition
  @param  AHideText    Hides the normal cell text
  @param  AReverse     Icons are assigned in reverse order.

  @seeAlso  TsCFIconSet
  @seeAlso  TsCFValueKind
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Writes the conditional format "icon set" in which cell values are illustrated
  by means of four icons.

  @param  ARange       Range of the cells to be formatted
  @param  AIconSet     Selection of the icon set to be used.
  @param  AValueKind1  Condition for the cell values to be displayed with the first icon in the icon set
  @param  AValue1      Value to be compared with in the AValueKind1 condition
  @param  AValueKind2  Condition for the cell values to be displayed with the second icon in the icon set
  @param  AValue2      Value to be compared with in the AValueKind2 condition
  @param  AValueKind3  Condition for the cell values to be displayed with the third icon in the icon set
  @param  AValue3      Value to be compared with in the AValueKind3 condition
  @param  AHideText    Hides the normal cell text
  @param  AReverse     Icons are assigned in reverse order.

  @seeAlso  TsCFIconSet
  @seeAlso  TsCFValueKind
-------------------------------------------------------------------------------}
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


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "icon set" in which cell values are illustrated
  by means of five icons.

  @param  ARange       Range of the cells to be formatted
  @param  AIconSet     Selection of the icon set to be used.
  @param  AValueKind1  Condition for the cell values to be displayed with the first icon in the icon set
  @param  AValue1      Value to be compared with in the AValueKind1 condition
  @param  AValueKind2  Condition for the cell values to be displayed with the second icon in the icon set
  @param  AValue2      Value to be compared with in the AValueKind2 condition
  @param  AValueKind3  Condition for the cell values to be displayed with the third icon in the icon set
  @param  AValue3      Value to be compared with in the AValueKind3 condition
  @param  AValueKind4  Condition for the cell values to be displayed with the fourth icon in the icon set
  @param  AValue4      Value to be compared with in the AValueKind4 condition
  @param  AHideText    Hides the normal cell text
  @param  AReverse     Icons are assigned in reverse order.

  @seeAlso  TsCFIconSet
  @seeAlso  TsCFValueKind
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Returns the conditional format record at the specified index in the workbook's
  conditional format list.
-------------------------------------------------------------------------------}
function TsWorkbook.GetConditionalFormat(AIndex: Integer): TsConditionalFormat;
begin
  Result := FConditionalFormatList[AIndex] as TsConditionalFormat;
end;

{@@ ----------------------------------------------------------------------------
  Returns the count of conditional formats available in the the workbook.
-------------------------------------------------------------------------------}
function TsWorkbook.GetNumConditionalFormats: Integer;
begin
  Result := FConditionalFormatList.Count;
end;

