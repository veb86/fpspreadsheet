{ Included by fpspreadsheet.pas }

{ Returns the count of conditional format items }
function TsWorksheet.GetConditionalFormatCount: Integer;
begin
  Result := FConditionalFormats.Count;
end;

{@@ ----------------------------------------------------------------------------
  Returns the conditional format item stored in the CF list at the specified
  index.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadConditionalFormat(AIndex: Integer): TsConditionalFormat;
begin
  Result := TsConditionalFormat(FConditionalFormats[AIndex]);
end;

{@@ ----------------------------------------------------------------------------
  Creates a conditional format item for the cells given by ARange.
  The condition specified must not require parameters, e.g. cfcEmpty
-------------------------------------------------------------------------------}
function TsWorksheet.WriteConditionalCellFormat(ARange: TsCellRange;
  ACondition: TsCFCondition; ACellFormatIndex: Integer): Integer;
begin
  Result := FConditionalFormats.AddCellRule(ARange, ACondition,
    ACellFormatIndex);
end;

{@@ ----------------------------------------------------------------------------
  Creates a conditional format item for the cells given by ARange.
  The condition specified must require one parameter, e.g. cfcEqual,
  and the parameter must be specified as AParam.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteConditionalCellFormat(ARange: TsCellRange;
  ACondition: TsCFCondition; AParam: Variant; ACellFormatIndex: Integer): Integer;
begin
  Result := FConditionalFormats.AddCellRule(ARange, ACondition,
    AParam, ACellFormatIndex);
end;

{@@ ----------------------------------------------------------------------------
  Creates a conditional format item for the cells given by ARange.
  The condition specified must requored two parameters, e.g. cfcBetween,
  and the parameters must be specified as AParam1 and AParam2.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteConditionalCellFormat(ARange: TsCellRange;
  ACondition: TsCFCondition; AParam1, AParam2: Variant;
  ACellFormatIndex: Integer): Integer;
begin
  Result := FConditionalFormats.AddCellRule(ARange, ACondition,
    AParam1, AParam2, ACellFormatIndex);
end;

