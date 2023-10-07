{ Included by fpspreadsheet.pas }
{ Contains code for cell formatting }

{==============================================================================}
{  TsWorksheet code for format handling                                        }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Modifies the background parameters of the format record stored at the
  specified index.

  @param   AFormatIndex      Index of the format record to be changed
  @param   AStyle            Fill style ("pattern") to be used - see @link(TsFillStyle)
  @param   APatternColor     RGB value of the pattern color
  @param   ABackgroundColor  RGB value of the background color

  @returns Index of the new format record.

  @NOTE   When AStyle is fsSolidFill the color is defined by APatternColor, ABackgroundColor is ignored unless the APatternColor is not used (scTransparent).
-------------------------------------------------------------------------------}
function TsWorksheet.ChangeBackground(AFormatIndex: Integer; AStyle: TsFillStyle;
  APatternColor: TsColor = scTransparent;
  ABackgroundColor: TsColor = scTransparent): Integer;
var
  fmt: TsCellFormat;
begin
  fmt := Workbook.GetCellFormat(AFormatIndex);
  if (AStyle = fsNoFill) or
     ((APatternColor = scTransparent) and (ABackgroundColor = scTransparent))
  then
    Exclude(fmt.UsedFormattingFields, uffBackground)
  else
  begin
    Include(fmt.UsedFormattingFields, uffBackground);
    fmt.Background.Style := AStyle;
    fmt.Background.FgColor := APatternColor;
    if (AStyle = fsSolidFill) and (APatternColor = scTransparent) then
      fmt.Background.FgColor := ABackgroundColor
    else
      fmt.Background.BgColor := ABackgroundColor;
  end;
  Result := Workbook.AddCellFormat(fmt);
end;


{@@ ----------------------------------------------------------------------------
  Returns the background fill pattern and colors of a cell.

  @param    ACell  Pointer to the cell
  @returns  @link(TsFillPattern) record (or @link(EMPTY_FILL), if the cell does not have a filled background
-------------------------------------------------------------------------------}
function TsWorksheet.ReadBackground(ACell: PCell): TsFillPattern;
var
  fmt : PsCellFormat;
begin
  Result := EMPTY_FILL;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffBackground in fmt^.UsedFormattingFields) then
      Result := fmt^.Background;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Returns the background color of a cell as rbg value

  @param     ACell   Pointer to the cell
  @returns   Value containing the rgb bytes in little-endian order
-------------------------------------------------------------------------------}
function TsWorksheet.ReadBackgroundColor(ACell: PCell): TsColor;
begin
  Result := scTransparent;
  if ACell <> nil then
    Result := ReadBackgroundColor(ACell^.FormatIndex);
end;


{@@ ----------------------------------------------------------------------------
  Returns the background color stored at the specified index in the format
  list of the workkbok.

  @param     AFormatIndex    Index of the format record
  @returns   Value containing the rgb bytes in little-endian order
-------------------------------------------------------------------------------}
function TsWorksheet.ReadBackgroundColor(AFormatIndex: Integer): TsColor;
var
  fmt: PsCellFormat;
begin
  Result := scTransparent;
  if AFormatIndex > -1 then begin
    fmt := Workbook.GetPointerToCellFormat(AFormatIndex);
    if (uffBackground in fmt^.UsedFormattingFields) then
    begin
      if fmt^.Background.Style = fsSolidFill then
        Result := fmt^.Background.FgColor
      else
        Result := fmt^.Background.BgColor;
    end;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Returns the BiDi mode of the cell (right-to-left or left-to-right)
-------------------------------------------------------------------------------}
function TsWorksheet.ReadBiDiMode(ACell: PCell): TsBiDiMode;
var
  fmt: PsCellFormat;
begin
  Result := bdDefault;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffBiDi in fmt^.UsedFormattingFields) then
      Result := fmt^.BiDiMode;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Determines which borders are drawn around a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellBorders(ACell: PCell): TsCellBorders;
var
  fmt: PsCellFormat;
begin
  Result := [];
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffBorder in fmt^.UsedFormattingFields) then
      Result := fmt^.Border;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines the style of a particular cell border
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellBorderStyle(ACell: PCell;
  ABorder: TsCelLBorder): TsCellBorderStyle;
var
  fmt: PsCellFormat;
begin
  Result := DEFAULT_BORDERSTYLES[ABorder];
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := fmt^.BorderStyles[ABorder];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines all border styles of a given cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellBorderStyles(ACell: PCell): TsCellBorderStyles;
var
  fmt: PsCellFormat;
begin
  Result := DEFAULT_BORDERSTYLES;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := Fmt^.BorderStyles;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the format record that is assigned to a specified cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellFormat(ACell: PCell): TsCellFormat;
begin
  Result := Workbook.GetCellFormat(ACell^.FormatIndex);
end;

{@@ ----------------------------------------------------------------------------
  Returns the protection flags of the cell.

  @NOTE These flags are active only if sheet protection is active, i.e. soProtected in Worksheet.Options.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellProtection(ACell: PCell): TsCellProtections;
var
  fmt: PsCellFormat;
begin
  Result := DEFAULT_CELL_PROTECTION;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if fmt <> nil then
      Result := fmt^.Protection;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns whether the cell is not allowed to be printed.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadDoNotPrintCell(ACell: PCell): Boolean;
var
  fmt: PsCellFormat;
begin
  Result := false;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := uffDoNotPrint in fmt.UsedFormattingFields;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the horizontal alignment of a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadHorAlignment(ACell: PCell): TsHorAlignment;
var
  fmt: PsCellFormat;
begin
  Result := haDefault;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffHorAlign in fmt^.UsedFormattingFields) then
      Result := fmt^.HorAlignment;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Returns the text orientation of a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadTextRotation(ACell: PCell): TsTextRotation;
var
  fmt: PsCellFormat;
begin
  Result := trHorizontal;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffTextRotation in fmt^.UsedFormattingFields) then
      Result := fmt^.TextRotation;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Reads the set of used formatting fields of a cell.

  Each cell contains a set of "used formatting fields". Formatting is applied
  only if the corresponding element is contained in the set.

  @param    ACell   Pointer to the cell
  @returns  Set of elements used in formatting the cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadUsedFormatting(ACell: PCell): TsUsedFormattingFields;
var
  fmt: PsCellFormat;
begin
  if ACell = nil then
  begin
    Result := [];
    Exit;
  end;
  fmt := FWorkbook.GetPointerToCellFormat(ACell^.FormatIndex);
  Result := fmt^.UsedFormattingFields;
end;


{@@ ----------------------------------------------------------------------------
  Returns the vertical alignment of a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadVertAlignment(ACell: PCell): TsVertAlignment;
var
  fmt: PsCellFormat;
begin
  Result := vaDefault;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffVertAlign in fmt^.UsedFormattingFields) then
      Result := fmt^.VertAlignment;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Returns whether a specific cell support word-wrapping.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadWordwrap(ACell: PCell): boolean;
var
  fmt: PsCellFormat;
begin
  Result := false;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := uffWordwrap in fmt^.UsedFormattingFields;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Defines a background pattern for a cell

  @param   ARow              Row index of the cell
  @param   ACol              Column index of the cell
  @param   AFillStyle        Fill style to be used - see @link(TsFillStyle)
  @param   APatternColor     RGB value of the pattern color
  @param   ABackgroundColor  RGB value of the background color
  @returns Pointer to cell

  @NOTE    When AStyle is fsSolidFill the color is defined by APatternColor, ABackgroundColor is ignored unless the APatternColor is not used (scTransparent).
  @NOTE   Is replaced by uniform fill if WriteBackgroundColor is called later.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBackground(ARow, ACol: Cardinal; AStyle: TsFillStyle;
  APatternColor, ABackgroundColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBackground(Result, AStyle, APatternColor, ABackgroundColor);
end;


{@@ ----------------------------------------------------------------------------
  Defines a background pattern for a cell

  @param  ACell             Pointer to the cell
  @param  AStyle            Fill style ("pattern") to be used - see @link(TsFillStyle)
  @param  APatternColor     RGB value of the pattern color
  @param  ABackgroundColor  RGB value of the background color

  @NOTE   When AStyle is fsSolidFill the color is defined by APatternColor, ABackgroundColor is ignored unless the APatternColor is not used (scTransparent).
  @NOTE   Is replaced by uniform fill if WriteBackgroundColor is called later.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBackground(ACell: PCell; AStyle: TsFillStyle;
  APatternColor: TsColor = scTransparent; ABackgroundColor: TsColor = scTransparent);
var
  idx: Integer;
begin
  if ACell <> nil then begin
    idx := ACell^.FormatIndex;
    ACell^.FormatIndex := ChangeBackground(idx, AStyle, APatternColor, ABackgroundColor);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Sets a uniform background color of a cell.

  @param   ARow       Row index of the cell
  @param   ACol       Column index of the cell
  @param   AColor     RGB value of the new background color. Use the value "scTransparent" to clear an existing background color.
  @returns Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBackgroundColor(ARow, ACol: Cardinal;
  AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBackgroundColor(Result, AColor);
end;


{@@ ----------------------------------------------------------------------------
  Sets a uniform background color of a cell.

  @param  ACell      Pointer to cell
  @param  AColor     RGB value of the new background color. Use the value "scTransparent" to clear an existing background color.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBackgroundColor(ACell: PCell; AColor: TsColor);
begin
  if ACell <> nil then begin
    if AColor = scTransparent then
      WriteBackground(ACell, fsNoFill)
    else
      WriteBackground(ACell, fsSolidFill, AColor, AColor);
  end;
end;


function TsWorksheet.WriteBiDiMode(ARow, ACol: Cardinal; AValue: TsBiDiMode): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBiDiMode(Result, AValue);
end;


procedure TsWorksheet.WriteBiDiMode(ACell: PCell; AValue: TsBiDiMode);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  fmt.BiDiMode := AValue;
  if AValue <> bdDefault then
    Include(fmt.UsedFormattingFields, uffBiDi)
  else
    Exclude(fmt.UsedFormattingFields, uffBiDi);
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;


{@@ ----------------------------------------------------------------------------
  Sets the color of a cell border line.
  Note: the border must be included in Borders set in order to be shown!

  @param   ARow       Row index of the cell
  @param   ACol       Column index of the cell
  @param   ABorder    Indicates to which border (left/top etc) this color is to be applied
  @param   AColor     RGB value of the new border color
  @returns Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderColor(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderColor(Result, ABorder, AColor);
end;


{@@ ----------------------------------------------------------------------------
  Sets the color of a cell border line.
  Note: the border must be included in Borders set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Indicates to which border (left/top etc) this color is to be applied
  @param  AColor     RGB value of the new border color
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderColor(ACell: PCell; ABorder: TsCellBorder;
  AColor: TsColor);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder].Color := AColor;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Sets the linestyle of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param   ARow       Row index of the cell
  @param   ACol       Column index of the cell
  @param   ABorder    Indicates to which border (left/top etc) this color is to be applied
  @param   ALineStyle Identifier of the new line style to be applied.
  @returns Pointer to cell

  @seeAlso TsLineStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderLineStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderLineStyle(Result, ABorder, ALineStyle);
end;


{@@ ----------------------------------------------------------------------------
  Sets the linestyle of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param   ACell      Pointer to cell
  @param   ABorder    Indicates to which border (left/top etc) this color is to be applied
  @param   ALineStyle Identifier of the new line style to be applied.

  @seeAlso TsLineStyle
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderLineStyle(ACell: PCell;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder].LineStyle := ALineStyle;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Shows the cell borders included in the set ABorders. No border lines are drawn
  for those not included.

  The borders are drawn using the "BorderStyles" assigned to the cell.

  @param   ARow      Row index of the cell
  @param   ACol      Column index of the cell
  @param   ABorders  Set with elements to identify the border(s) to will be shown
  @returns Pointer to cell
  @seeAlso TsCellBorder
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorders(ARow, ACol: Cardinal; ABorders: TsCellBorders): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorders(Result, ABorders);
end;


{@@ ----------------------------------------------------------------------------
  Shows the cell borders included in the set ABorders. No border lines are drawn
  for those not included.

  The borders are drawn using the "BorderStyles" assigned to the cell.

  @param   ACell     Pointer to cell
  @param   ABorders  Set with elements to identify the border(s) to will be shown
  @seeAlso TsCellBorder
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorders(ACell: PCell; ABorders: TsCellBorders);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    if ABorders = [] then
      Exclude(fmt.UsedFormattingFields, uffBorder)
    else
      Include(fmt.UsedFormattingFields, uffBorder);
    fmt.Border := ABorders;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Sets the style of a cell border, i.e. line style and line color.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param   ARow       Row index of the cell considered
  @param   ACol       Column index of the cell considered
  @param   ABorder    Identifies the border to be modified (left/top/right/bottom)
  @param   AStyle     @link(TsCellBorderStyle) record with parameters controlling how the border line is drawn (line style, line color)
  @returns Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AStyle: TsCellBorderStyle): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyle(Result, ABorder, AStyle);
end;


{@@ ----------------------------------------------------------------------------
  Sets the style of a cell border, i.e. line style and line color.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Identifies the border to be modified (left/top/right/bottom)
  @param  AStyle     @link(TsCellBorderStyle) record with parameters controlling how the border line is drawn (line style, line color)
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
  AStyle: TsCellBorderStyle);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder] := AStyle;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Sets line style and line color of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param   ARow       Row index of the considered cell
  @param   ACol       Column index of the considered cell
  @param   ABorder    Identifier of the border to be modified
  @param   ALineStyle Identifier for the new line style of the border
  @param   AColor     RGB value of the border line color
  @returns Pointer to cell

  @seeAlso WriteBorderStyles
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle; AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyle(Result, ABorder, ALineStyle, AColor);
end;


{@@ ----------------------------------------------------------------------------
  Sets line style and line color of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Identifier of the border to be modified
  @param  ALineStyle Identifier for the new line style of the border
  @param  AColor     RGB value of the color of the border line

  @seeAlso WriteBorderStyles
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
  ALineStyle: TsLineStyle; AColor: TsColor);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder].LineStyle := ALineStyle;
    fmt.BorderStyles[ABorder].Color := AColor;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Sets the style of all cell border of a cell, i.e. line style and line color.
  Note: Only those borders included in the "Borders" set are shown!

  @param   ARow     Row index of the considered cell
  @param   ACol     Column index of the considered cell
  @param   AStyles  Array of @link(CellBorderStyle) records for each cell border.
  @returns Pointer to cell

  @seeAlso WriteBorderStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderStyles(ARow, ACol: Cardinal;
  const AStyles: TsCellBorderStyles): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyles(Result, AStyles);
end;


{@@ ----------------------------------------------------------------------------
  Sets the style of all cell border of a cell, i.e. line style and line color.
  Note: Only those borders included in the "Borders" set are shown!

  @param  ACell    Pointer to cell
  @param  ACol     Column index of the considered cell
  @param  AStyles  Array of @link(CellBorderStyle) records for each cell border.

  @seeAlso WriteBorderStyle
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderStyles(ACell: PCell;
  const AStyles: TsCellBorderStyles);
var
  b: TsCellBorder;
  fmt: TsCellFormat;
begin
  if Assigned(ACell) then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    for b in TsCellBorder do fmt.BorderStyles[b] := AStyles[b];
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Assigns a complete cell format record to a cell

  @param ACell        Pointer to the cell to be modified
  @param ACellFormat  Cell format record to be used by the cell

  @seeAlso TsCellFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteCellFormat(ACell: PCell;
  const ACellFormat: TsCellFormat);
var
  idx: Integer;
begin
  idx := Workbook.AddCellFormat(ACellFormat);
  WriteCellFormatIndex(ACell, idx);
end;


{@@ ----------------------------------------------------------------------------
  Formats a cell to the cell format stored at the specified index in the
  workbook's cell format list.

  @param  ACell       Pointer to the cell to be formatted
  @param  AIndex      Index of the cell format record to be used by the cell

  @seeAlso TsCellFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteCellFormatIndex(ACell: PCell; AIndex: Integer);
begin
  if AIndex >= Workbook.GetNumCellFormats then
    raise EFpSpreadsheet.Create('[Worksheet.WriteCellFormat] Invalid cell format index.');

  // The default format index is 0, but it could also be refered to by -1
  if AIndex < 0 then AIndex := 0;
  if Assigned(ACell) then begin
    ACell^.FormatIndex := AIndex;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Defines how the cell at the specified row and column is protected: lock
  cell modification and/or hide formulas. Note that this is activated only after
  enabling worksheet protection (worksheet.Protect(true)).

  @NOTE FPSpreadsheet does not enforce these actions. They are only written to the file for the Office application.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteCellProtection(ARow, ACol: Cardinal;
  AValue: TsCellProtections): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteCellProtection(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Defines how the specified cell is protected: lock cell modification and/or hide formulas.
  Note that this is activated only after enabling worksheet protection (worksheet.Protect(true)).

  @NOTE FPSpreadsheet does not enforce these actions. They are only written to the file for the Office application.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteCellProtection(ACell: PCell;
  AValue: TsCellProtections);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  fmt.Protection := AValue;
  if AValue <> DEFAULT_CELL_PROTECTION then
    Include(fmt.UsedFormattingFields, uffProtection)
  else
    Exclude(fmt.UsedFormattingFields, uffProtection);
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

function TsWorksheet.WriteDoNotPrintCell(ARow, ACol: Cardinal;
  AValue: boolean): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDoNotPrintCell(Result, AValue);
end;

procedure TsWorksheet.WriteDoNotPrintCell(ACell: PCell; AValue: Boolean);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if AValue then
    Include(fmt.UsedFormattingfields, uffDoNotPrint)
  else
    Exclude(fmt.UsedFormattingFields, uffDoNotPrint);
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;


{@@ ----------------------------------------------------------------------------
  Defines the horizontal alignment of text in a cell.

  @param ARow    Row index of the cell considered
  @param ACol    Column index of the cell considered
  @param AValue  Parameter for horizontal text alignment (haDefault, vaLeft, haCenter, haRight). By default, texts are left-aligned, numbers and dates are right-aligned.
  @returns Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteHorAlignment(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Defines the horizontal alignment of text in a cell.

  @param ACell   Pointer to the cell considered
  @param AValue  Parameter for horizontal text alignment (haDefault, vaLeft, haCenter, haRight). By default, texts are left-aligned, numbers and dates are right-aligned.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteHorAlignment(ACell: PCell; AValue: TsHorAlignment);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if AValue = haDefault then
    Exclude(fmt.UsedFormattingFields, uffHorAlign)
  else
    Include(fmt.UsedFormattingFields, uffHorAlign);
  fmt.HorAlignment := AValue;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Adds text rotation to the formatting of a cell

  @param   ARow      The row of the cell
  @param   ACol      The column of the cell
  @param   ARotation How to rotate the text
  @returns Pointer to cell

  @seeAlso TsTextRotation
-------------------------------------------------------------------------------}
function TsWorksheet.WriteTextRotation(ARow, ACol: Cardinal;
  ARotation: TsTextRotation): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteTextRotation(Result, ARotation);
end;

{@@ ----------------------------------------------------------------------------
  Adds text rotation to the formatting of a cell

  @param    ACell      Pointer to the cell
  @param    ARotation  How to rotate the text

  @seeAlso  TsTextRotation
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteTextRotation(ACell: PCell; ARotation: TsTextRotation);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  Include(fmt.UsedFormattingFields, uffTextRotation);
  fmt.TextRotation := ARotation;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedFont(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Directly modifies the used formatting fields of a cell.
  Only formatting corresponding to items included in this set is executed.

  @param   ARow            The row of the cell
  @param   ACol            The column of the cell
  @param   AUsedFormatting set of the used formatting fields
  @returns Pointer to the (existing or created) cell

  @seeAlso TsUsedFormattingFields
  @seeAlso TCell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteUsedFormatting(ARow, ACol: Cardinal;
  AUsedFormatting: TsUsedFormattingFields): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteUsedFormatting(Result, AUsedFormatting);
end;

{@@ ----------------------------------------------------------------------------
  Directly modifies the used formatting fields of an existing cell.
  Only formatting corresponding to items included in this set is executed.

  @param   ACell            Pointer to the cell to be modified
  @param   AUsedFormatting  Set of the used formatting fields

  @seeAlso TsUsedFormattingFields
  @seeAlso TCell
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteUsedFormatting(ACell: PCell;
  AUsedFormatting: TsUsedFormattingFields);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  fmt.UsedFormattingFields := AUsedFormatting;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Defines the vertical alignment of text in a cell.

  @param ARow     Row index of the cell considered
  @param ACol     Column index of the cell considered
  @param AValue   Parameter for vertical text alignment (vaDefault, vaTop, vaCenter, vaBottom). By default, texts are bottom-aligned.
  @returns        Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteVertAlignment(ARow, ACol: Cardinal;
  AValue: TsVertAlignment): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteVertAlignment(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Defines the vertical alignment of text in a cell.

  @param ACell   Poiner to the cell considered
  @param AValue  Parameter for vertical text alignment (vaDefault, vaTop, vaCenter, vaBottom). By default, texts are bottom-aligned.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteVertAlignment(ACell: PCell; AValue: TsVertAlignment);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if AValue = vaDefault then
    Exclude(fmt.UsedFormattingFields, uffVertAlign)
  else
    Include(fmt.UsedFormattingFields, uffVertAlign);
  fmt.VertAlignment := AValue;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Enables or disables the word-wrapping feature for a cell.

  @param   ARow    Row index of the cell considered
  @param   ACol    Column index of the cell considered
  @param   AValue  @true = word-wrapping enabled, @false = disabled.
  @returns Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteWordwrap(ARow, ACol: Cardinal; AValue: boolean): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteWordWrap(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Enables or disables the word-wrapping feature for a cell.

  @param ACel    Pointer to the cell considered
  @param AValue  @true = word-wrapping enabled, @false = disabled.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteWordwrap(ACell: PCell; AValue: boolean);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if AValue then
    Include(fmt.UsedFormattingFields, uffWordwrap)
  else
    Exclude(fmt.UsedFormattingFields, uffWordwrap);
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;


{==============================================================================}
{  TsWorkbook code for format handling                                         }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Adds the specified format record to the internal list and returns the index
  in the list. If the record had already been added before the function only
  returns the index.
-------------------------------------------------------------------------------}
function TsWorkbook.AddCellFormat(const AValue: TsCellFormat): Integer;
begin
  Result := FCellFormatList.Add(AValue);
end;


{@@ ----------------------------------------------------------------------------
  Returns the contents of the format record with the specified index.
-------------------------------------------------------------------------------}
function TsWorkbook.GetCellFormat(AIndex: Integer): TsCellFormat;
begin
  Result := FCellFormatList.Items[AIndex]^;
end;


{@@ ----------------------------------------------------------------------------
  Returns a string describing the cell format with the specified index.
-------------------------------------------------------------------------------}
function TsWorkbook.GetCellFormatAsString(AIndex: Integer): String;
var
  fmt: PsCellFormat;
  cb: TsCellBorder;
  s: String;
  numFmt: TsNumFormatParams;
begin
  Result := '';
  fmt := GetPointerToCellFormat(AIndex);
  if fmt = nil then
    exit;

  if (uffFont in fmt^.UsedFormattingFields) then
    Result := Format('%s; Font%d', [Result, fmt^.FontIndex]);
  if (uffBackground in fmt^.UsedFormattingFields) then begin
    Result := Format('%s; Bg %s', [Result, GetColorName(fmt^.Background.BgColor)]);
    Result := Format('%s; Fg %s', [Result, GetColorName(fmt^.Background.FgColor)]);
    Result := Format('%s; Pattern %s', [Result, GetEnumName(TypeInfo(TsFillStyle), ord(fmt^.Background.Style))]);
  end;
  if (uffHorAlign in fmt^.UsedFormattingfields) then
    Result := Format('%s; %s', [Result, GetEnumName(TypeInfo(TsHorAlignment), ord(fmt^.HorAlignment))]);
  if (uffVertAlign in fmt^.UsedFormattingFields) then
    Result := Format('%s; %s', [Result, GetEnumName(TypeInfo(TsVertAlignment), ord(fmt^.VertAlignment))]);
  if (uffWordwrap in fmt^.UsedFormattingFields) then
    Result := Format('%s; Word-wrap', [Result]);
  if (uffNumberFormat in fmt^.UsedFormattingFields) then
  begin
    numFmt := GetNumberFormat(fmt^.NumberFormatIndex);
    if numFmt <> nil then
      Result := Format('%s; %s (%s)', [Result,
        GetEnumName(TypeInfo(TsNumberFormat), ord(numFmt.NumFormat)),
        numFmt.NumFormatStr
      ])
    else
      Result := Format('%s; %s', [Result, 'nfGeneral']);
  end else
    Result := Format('%s; %s', [Result, 'nfGeneral']);
  if (uffBorder in fmt^.UsedFormattingFields) then
  begin
    s := '';
    for cb in fmt^.Border do
      if s = '' then s := GetEnumName(TypeInfo(TsCellBorder), ord(cb))
        else s := s + '+' + GetEnumName(TypeInfo(TsCellBorder), ord(cb));
    Result := Format('%s; %s', [Result, s]);
  end;
  if (uffBiDi in fmt^.UsedFormattingFields) then
    Result := Format('%s; %s', [Result, GetEnumName(TypeInfo(TsBiDiMode), ord(fmt^.BiDiMode))]);
  if Result <> '' then Delete(Result, 1, 2);
end;


{@@ ----------------------------------------------------------------------------
  Returns the count of format records used all over the workbook
-------------------------------------------------------------------------------}
function TsWorkbook.GetNumCellFormats: Integer;
begin
  Result := FCellFormatList.Count;
end;


{@@ ----------------------------------------------------------------------------
  Returns a pointer to the format record with the specified index
-------------------------------------------------------------------------------}
function TsWorkbook.GetPointerToCellFormat(AIndex: Integer): PsCellFormat;
begin
  if FCellFormatList.Count = 0 then
    raise Exception.Create('[TsWorkbook.GetPointerToCellFormat]: No format items.');

  if (AIndex < 0) or (AIndex >= FCellFormatList.Count) then
    AIndex := 0;  // 0 is default format
  Result := FCellFormatList.Items[AIndex];
end;


{@@ ----------------------------------------------------------------------------
  Removes all cell formats from the workbook.

  If AKeepDefaultFormat is @true then index 0 containing the default cell format
  is retained.

  Use carefully!
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveAllCellFormats(AKeepDefaultFormat: Boolean);
var
  i: Integer;
begin
  if AKeepDefaultFormat then
    for i := FCellFormatList.Count-1 downto 1 do
      FCellFormatList.Delete(i)
  else
    FCellFormatList.Clear;
end;

