{ Included by fpspreadsheet.pas }

{ Code for number format }

{==============================================================================}
{   TsWorksheet code for number format                                         }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Determines some number format attributes (decimal places, currency symbol) of
  a cell

  @param    ACell           Pointer to the cell under investigation
  @param    ADecimals       Number of decimal places that can be extracted from the formatting string, e.g. in case of '0.000' this would be 3.
  @param    ACurrencySymbol String representing the currency symbol extracted from the formatting string.

  @returns  @true if the the format string could be analyzed successfully, @false if not
-------------------------------------------------------------------------------}
function TsWorksheet.GetNumberFormatAttributes(ACell: PCell; out ADecimals: byte;
  out ACurrencySymbol: String): Boolean;
var
  parser: TsNumFormatParser;
  nf: TsNumberFormat;
  nfs: String;
begin
  Result := false;
  if ACell <> nil then
  begin
    ReadNumFormat(ACell, nf, nfs);
    parser := TsNumFormatParser.Create(nfs, FWorkbook.FormatSettings);
    try
      if parser.Status = psOK then
      begin
        nf := parser.NumFormat;
        if (nf = nfGeneral) and (ACell^.ContentType = cctNumber) then
        begin
          ADecimals := GetDisplayedDecimals(ACell);
          ACurrencySymbol := '';
        end else
        if IsDateTimeFormat(nf) then
        begin
          ADecimals := 2;
          ACurrencySymbol := '?';
        end
        else
        begin
          ADecimals := parser.Decimals;
          ACurrencySymbol := parser.CurrencySymbol;
        end;
        Result := true;
      end;
    finally
      parser.Free;
    end;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Returns the number format type and format string used in a specific cell
  
  @param  ACell          Pointer to the cell of interest
  @param  ANumFormat     Number format record to be used in this cell.
  @param  ANumFormatStr  Number format string to be used in this cell.
-------------------------------------------------------------------------------}
procedure TsWorksheet.ReadNumFormat(ACell: PCell; out ANumFormat: TsNumberFormat;
  out ANumFormatStr: String);
var
  fmt: PsCellFormat;
  numFmt: TsNumFormatParams;
begin
  ANumFormat := nfGeneral;
  ANumFormatStr := '';
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffNumberFormat in fmt^.UsedFormattingFields) then
    begin
      numFmt := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);
      if numFmt <> nil then
      begin
        ANumFormat := numFmt.NumFormat;
        ANumFormatStr := numFmt.NumFormatStr;
      end else
      begin
        ANumFormat := nfGeneral;
        ANumFormatStr := '';
      end;
    end;
  end;
end;


 {@@ ----------------------------------------------------------------------------
  Adds a date/time format to the formatting of a cell

  @param    ARow             The row index of the cell (zero-based) 
  @param    ACol             The column index of the cell (zero-based) 
  @param    ANumFormat       Identifier of the format to be applied (nfXXXX constant)
  @param    ANumFormatString Optional string of formatting codes. Is only considered if ANumberFormat is nfCustom.
  @returns  Pointer to the cell

  @seeAlso  TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDateTimeFormat(ARow, ACol: Cardinal;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTimeFormat(Result, ANumFormat, ANumFormatString);
end;


{@@ ----------------------------------------------------------------------------
  Adds a date/time format to the formatting of a cell

  @param   ACell             Pointer to the cell considered
  @param   ANumFormat        Identifier of the format to be applied (nxXXXX constant) 
  @param   ANumFormatString  Optional string of formatting codes. Is only considered if ANumberFormat is nfCustom.

  @seeAlso TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteDateTimeFormat(ACell: PCell;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = '');
var
  fmt: TsCellFormat;
  nfs: String;
  nfp: TsNumFormatParams;
  isTextFmt, wasTextFmt: Boolean;
  oldVal: String;
begin
  if ACell = nil then
    exit;

  if not ((ANumFormat in [nfGeneral, nfCustom]) or IsDateTimeFormat(ANumFormat)) then
    raise EFPSpreadsheet.Create('WriteDateTimeFormat can only be called with date/time formats.');

  isTextFmt := false;
  wasTextFmt := false;

  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  fmt.NumberFormat := ANumFormat;
  if (ANumFormat <> nfGeneral) then
  begin
    nfp := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
    wasTextFmt := IsTextFormat(nfp);
    oldval := ReadAsText(ACell);
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    if (ANumFormatString = '') then
      nfs := BuildDateTimeFormatString(ANumFormat, Workbook.FormatSettings)
    else
      nfs := ANumFormatString;
    isTextFmt := (nfs = '@');
  end else
  begin
    Exclude(fmt.UsedFormattingFields, uffNumberFormat);
    fmt.NumberFormatStr := '';
  end;
  fmt.NumberFormat := ANumFormat;
  fmt.NumberFormatStr := nfs;
  fmt.NumberFormatIndex := Workbook.AddNumberFormat(nfs);
  ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt);

  if isTextFmt  then
    WriteText(ACell, oldval)
  else
  if wasTextFmt then
    WriteCellValueAsString(ACell, ACell^.UTF8StringValue);

  ChangedCell(ACell^.Row, ACell^.Col);
end;


{@@ ----------------------------------------------------------------------------
  Formats the number in a cell to show a given count of decimal places.
  Is ignored for non-decimal formats (such as most date/time formats).

  @param    ARow       Row index of the cell considered (zero-based) 
  @param    ACol       Column index of the cell considered (zero-based) 
  @param    ADecimals  Number of decimal places to be displayed 
  @returns  Pointer to the cell
  @seeAlso  TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDecimals(ARow, ACol: Cardinal; ADecimals: Byte): PCell;
begin
  Result := FindCell(ARow, ACol);
  WriteDecimals(Result, ADecimals);
end;


{@@ ----------------------------------------------------------------------------
  Formats the numeric value in a cell to show a given count of decimal places.
  Is ignored for non-decimal formats (such as most date/time formats).

  @param    ACell      Pointer to the cell considered
  @param    ADecimals  Number of decimal places to be displayed
  @seeAlso  TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteDecimals(ACell: PCell; ADecimals: Byte);
var
  parser: TsNumFormatParser;
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams;
  numFmtStr: String;
begin
  if (ACell = nil) or (ACell^.ContentType <> cctNumber) then
    exit;

  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  numFmt := FWorkbook.GetNumberFormat(fmt.NumberFormatIndex);
  if numFmt <> nil then
    numFmtStr := numFmt.NumFormatStr
  else
    numFmtStr := '0.00';
  parser := TsNumFormatParser.Create(numFmtStr, Workbook.FormatSettings);
  try
    parser.Decimals := ADecimals;
    numFmtStr := parser.FormatString;
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(numFmtStr);
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  finally
    parser.Free;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Formats a number as a fraction

  @param    ARow             Row index of the cell
  @param    ACol             Column index of the cell
  @param    ANumFormat       Identifier of the format to be applied. Must be either nfFraction or nfMixedFraction
  @param    ANumeratorDigts  Count of numerator digits
  @param    ADenominatorDigits Count of denominator digits
  @returns  Pointer to the cell

  @seeAlso  TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFractionFormat(ARow, ACol: Cardinal;
  AMixedFraction: Boolean; ANumeratorDigits, ADenominatorDigits: Integer): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteFractionFormat(Result, AMixedFraction, ANumeratorDigits, ADenominatorDigits);
end;

{@@ ----------------------------------------------------------------------------
  Formats a number as a fraction

  @param   ACell              Pointer to the cell to be formatted
  @param   ANumFormat         Identifier of the format to be applied. Must be either nfFraction or nfMixedFraction
  @param   ANumeratorDigts    Count of numerator digits
  @param   ADenominatorDigits Count of denominator digits

  @seeAlso TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteFractionFormat(ACell: PCell;
  AMixedFraction: Boolean; ANumeratorDigits, ADenominatorDigits: Integer);
var
  fmt: TsCellFormat;
  nfs: String;
begin
  if ACell = nil then
    exit;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  nfs := BuildFractionFormatString(AMixedFraction, ANumeratorDigits, ADenominatorDigits);
  fmt.NumberFormatIndex := Workbook.AddNumberFormat(nfs);
  Include(fmt.UsedFormattingFields, uffNumberFormat);
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedCell(ACell^.Row, ACell^.Col);
end;


{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param   ARow            The row of the cell
  @param   ACol            The column of the cell
  @param   ANumFormat      Identifier of the format to be applied
  @param   ADecimals       Number of decimal places
  @param   ACurrencySymbol optional currency symbol in case of nfCurrency
  @param   APosCurrFormat  optional identifier for positive currencies
  @param   ANegCurrFormat  optional identifier for negative currencies
  @returns Pointer to the cell

  @seeAlso TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumFormat: TsNumberFormat; ADecimals: Integer; ACurrencySymbol: String = '';
  APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumberFormat(Result, ANumFormat, ADecimals, ACurrencySymbol,
    APosCurrFormat, ANegCurrFormat);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param   ARow            The row of the cell
  @param   ACol            The column of the cell
  @param   ANumFormat      Identifier of the format to be applied
  @param   ADecimals       Number of decimal places
  @param   ACurrencySymbol optional currency symbol in case of nfCurrency
  @param   APosCurrFormat  optional identifier for positive currencies
  @param   ANegCurrFormat  optional identifier for negative currencies

  @seeAlso TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteNumberFormat(ACell: PCell;
  ANumFormat: TsNumberFormat; ADecimals: Integer; ACurrencySymbol: String = '';
  APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1);
var
  fmt: TsCellFormat;
  fmtStr: String;
  nfp: TsNumFormatParams;
  wasTextFmt: Boolean;
begin
  if ACell = nil then
    exit;

  wasTextFmt := false;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  fmt.NumberFormat := ANumFormat;
  if ANumFormat <> nfGeneral then
  begin
    nfp := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
    wasTextFmt := IsTextFormat(nfp);
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    if IsCurrencyFormat(ANumFormat) then
    begin
      RegisterCurrency(ACurrencySymbol);
      fmtStr := BuildCurrencyFormatString(ANumFormat, Workbook.FormatSettings,
        ADecimals, APosCurrFormat, ANegCurrFormat, ACurrencySymbol);
    end else
      fmtStr := BuildNumberFormatString(ANumFormat,
        Workbook.FormatSettings, ADecimals);
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(fmtStr);
  end else
  begin
    Exclude(fmt.UsedFormattingFields, uffNumberFormat);
    fmt.NumberFormatIndex := -1;
  end;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  if wasTextFmt then
    WriteCellValueAsString(ACell, ACell^.UTF8StringValue);

  ChangedCell(ACell^.Row, ACell^.Col);
end;


{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param    ARow             The row of the cell
  @param    ACol             The column of the cell
  @param    ANumFormat       Identifier of the format to be applied
  @param    ANumFormatString Optional string of formatting codes. Is only considered if ANumberFormat is nfCustom.
  @returns  Pointer to the cell

  @seeAlso  TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumberFormat(Result, ANumFormat, ANumFormatString);
end;


{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param  ACell            Pointer to the cell considered
  @param  ANumFormat       Identifier of the format to be applied
  @param  ANumFormatString Optional string of formatting codes. Is only considered if ANumberFormat is nfCustom.

  @seeAlso    TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteNumberFormat(ACell: PCell;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = '');
var
  fmt: TsCellFormat;
  fmtStr: String;
  nfp: TsNumFormatParams;
  oldval: String;
  isTextFmt, wasTextFmt: Boolean;
begin
  if ACell = nil then
    exit;

  isTextFmt := false;
  wasTextFmt := false;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);

  if ANumFormat <> nfGeneral then begin
    nfp := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
    wasTextFmt := IsTextFormat(nfp);
    oldval := ReadAsText(ACell);
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    if (ANumFormatString = '') then
      fmtStr := BuildNumberFormatString(ANumFormat, Workbook.FormatSettings)
    else
      fmtStr := ANumFormatString;
    isTextFmt := (fmtstr = '@');
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(fmtStr);
  end else begin
    Exclude(fmt.UsedFormattingFields, uffNumberFormat);
    fmt.NumberFormatIndex := -1;
  end;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  if isTextFmt  then
    WriteText(ACell, oldval)
  else
  if wasTextFmt then
    WriteCellValueAsString(ACell, ACell^.UTF8StringValue);

  ChangedCell(ACell^.Row, ACell^.Col);
end;



{==============================================================================}
{   TsWorkbook code for number format                                          }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Adds a number format to the internal list. Returns the list index if already
  present, or creates a new format item and returns its index.
-------------------------------------------------------------------------------}
function TsWorkbook.AddNumberFormat(AFormatStr: String): Integer;
begin
  if AFormatStr = '' then
    Result := -1  // General number format is not stored
  else
    Result := TsNumFormatList(FNumFormatList).AddFormat(AFormatStr);
end;


{@@ ----------------------------------------------------------------------------
  Returns the parameters of the number format stored in the NumFormatList at the
  specified index.
  "General" number format is returned as nil.
-------------------------------------------------------------------------------}
function TsWorkbook.GetNumberFormat(AIndex: Integer): TsNumFormatParams;
begin
  if (AIndex >= 0) and (AIndex < FNumFormatList.Count) then
    Result := TsNumFormatParams(FNumFormatList.Items[AIndex])
  else
    Result := nil;
end;


{@@ ----------------------------------------------------------------------------
  Returns the count of number format records stored in the NumFormatList
-------------------------------------------------------------------------------}
function TsWorkbook.GetNumberFormatCount: Integer;
begin
  Result := FNumFormatList.Count;
end;


{@@ ----------------------------------------------------------------------------
  Removes all numberformats
  Use carefully!
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveAllNumberFormats;
var
  i: Integer;
  nfp: TsNumFormatParams;
begin
  for i:= FEmbeddedObjList.Count-1 downto 0 do begin
    nfp := TsNumFormatParams(FNumFormatList[i]);
    FNumFormatList.Delete(i);
    nfp.Free;
  end;
end;


