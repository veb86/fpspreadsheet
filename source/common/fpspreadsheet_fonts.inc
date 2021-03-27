{ Included by fpspreadsheet.pas }

{ Code for font handling }

{==============================================================================}
{   TsWorksheet code for fonts                                                 }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Determines the font used by a specified cell. Returns the workbook's default
  font if the cell does not exist.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellFont(ACell: PCell): TsFont;
var
  fmt: PsCellFormat;
begin
  Result := nil;
  if ACell <> nil then begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := Workbook.GetFont(fmt^.FontIndex);
  end;
  if Result = nil then
    Result := Workbook.GetDefaultFont;
end;

{@@ ----------------------------------------------------------------------------
  Determines the index of the font used by a specified cell, referring to the
  workbooks font list. Returns 0 (the default font index) if the cell does not
  exist.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellFontIndex(ACell: PCell): Integer;
var
  fmt: PsCellFormat;
begin
  Result := DEFAULT_FONTINDEX;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := fmt^.FontIndex;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Adds font specification to the formatting of a cell. Looks in the workbook's
  FontList and creates an new entry if the font is not used so far. Returns the
  index of the font in the font list.

  @param   ARow        The row of the cell
  @param   ACol        The column of the cell
  @param   AFontName   Name of the font
  @param   AFontSize   Size of the font, in points
  @param   AFontStyle  Set with font style attributes (don't use those of unit "graphics" !)
  @param   AFontColor  RGB value of the font's color
  @param   APosition   Specifies sub- or superscript text
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFont(ARow, ACol: Cardinal; const AFontName: String;
  AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor;
  APosition: TsFontPosition = fpNormal): Integer;
begin
  Result := WriteFont(GetCell(ARow, ACol), AFontName, AFontSize, AFontStyle,
    AFontColor, APosition);
end;


{@@ ----------------------------------------------------------------------------
  Adds font specification to the formatting of a cell. Looks in the workbook's
  FontList and creates an new entry if the font is not used so far. Returns the
  index of the font in the font list.

  @param   ACell       Pointer to the cell considered
  @param   AFontName   Name of the font
  @param   AFontSize   Size of the font, in points
  @param   AFontStyle  Set with font style attributes (don't use those of unit "graphics" !)
  @param   AFontColor  RGB value of the font's color
  @param   APosition   Specified subscript or superscript text.
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFont(ACell: PCell; const AFontName: String;
  AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor;
  APosition: TsFontPosition = fpNormal): Integer;
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
  begin
    Result := -1;
    Exit;
  end;

  Result := FWorkbook.FindFont(AFontName, AFontSize, AFontStyle, AFontColor, APosition);
  if Result = -1 then
    result := FWorkbook.AddFont(AFontName, AFontSize, AFontStyle, AFontColor, APosition);

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  Include(fmt.UsedFormattingFields, uffFont);
  fmt.FontIndex := Result;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedFont(ACell^.Row, ACell^.Col);
end;


{@@ ----------------------------------------------------------------------------
  Applies a font to the formatting of a cell. The font is determined by its
  index in the workbook's font list:

  @param   ARow        The row of the cell
  @param   ACol        The column of the cell
  @param   AFontIndex  Index of the font in the workbook's font list
  @returns Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFont(ARow, ACol: Cardinal; AFontIndex: Integer): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteFont(Result, AFontIndex);
end;


{@@ ----------------------------------------------------------------------------
  Applies a font to the formatting of a cell. The font is determined by its
  index in the workbook's font list:

  @param  ACell       Pointer to the cell considered
  @param  AFontIndex  Index of the font in the workbook's font list
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteFont(ACell: PCell; AFontIndex: Integer);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;

  if (AFontIndex < 0) or (AFontIndex >= Workbook.GetFontCount) then
    raise EFPSpreadsheet.Create(rsInvalidFontIndex);

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  Include(fmt.UsedFormattingFields, uffFont);
  fmt.FontIndex := AFontIndex;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedFont(ACell^.Row, ACell^.Col);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the text color used in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param   ARow        The row of the cell
  @param   ACol        The column of the cell
  @param   AFontColor  RGB value of the new text color
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontColor(ARow, ACol: Cardinal; AFontColor: TsColor): Integer;
begin
  Result := WriteFontColor(GetCell(ARow, ACol), AFontColor);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the text color used in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param   ACell       Pointer to the cell
  @param   AFontColor  RGB value of the new text color
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontColor(ACell: PCell; AFontColor: TsColor): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  Result := WriteFont(ACell, fnt.FontName, fnt.Size, fnt.Style, AFontColor);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the font used in formatting of a cell considering only the font face
  and leaving font size, style and color unchanged. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param   ARow        The row of the cell
  @param   ACol        The column of the cell
  @param   AFontName   Name of the new font to be used
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontName(ARow, ACol: Cardinal; AFontName: String): Integer;
begin
  result := WriteFontName(GetCell(ARow, ACol), AFontName);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the font used in formatting of a cell considering only the font face
  and leaving font size, style and color unchanged. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param   ACell       Pointer to the cell
  @param   AFontName   Name of the new font to be used
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontName(ACell: PCell; AFontName: String): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  result := WriteFont(ACell, AFontName, fnt.Size, fnt.Style, fnt.Color);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the font size in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param   ARow        The row of the cell
  @param   ACol        The column of the cell
  @param   ASize       Size of the font to be used (in points).
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontSize(ARow, ACol: Cardinal; ASize: Single): Integer;
begin
  Result := WriteFontSize(GetCell(ARow, ACol), ASize);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the font size in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param   ACell       Pointer to the cell
  @param   ASize       Size of the font to be used (in points).
  @returns Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontSize(ACell: PCell; ASize: Single): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  Result := WriteFont(ACell, fnt.FontName, ASize, fnt.Style, fnt.Color);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the font style (bold, italic, etc) in formatting of a cell.
  Looks in the workbook's font list if this modified font has already been used.
  If not a new font entry is created.
  Returns the index of this font in the font list.

  @param   ARow        The row of the cell
  @param   ACol        The column of the cell
  @param   AStyle      New font style to be used
  @returns Index of the font in the workbook's font list.

  @seeAlso TsFontStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontStyle(ARow, ACol: Cardinal;
  AStyle: TsFontStyles): Integer;
begin
  Result := WriteFontStyle(GetCell(ARow, ACol), AStyle);
end;


{@@ ----------------------------------------------------------------------------
  Replaces the font style (bold, italic, etc) in formatting of a cell.
  Looks in the workbook's font list if this modified font has already been used.
  If not a new font entry is created.
  Returns the index of this font in the font list.

  @param   ACell       Pointer to the cell considered
  @param   AStyle      New font style to be used
  @returns Index of the font in the workbook's font list.

  @seeAlso TsFontStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontStyle(ACell: PCell; AStyle: TsFontStyles): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := -1;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  Result := WriteFont(ACell, fnt.FontName, fnt.Size, AStyle, fnt.Color);
end;



{==============================================================================}
{   TsWorkbook code for fonts                                                  }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Adds a font to the font list. Returns the index in the font list.

  @param AFontName  Name of the font (like 'Arial')
  @param ASize      Size of the font in points
  @param AStyle     Style of the font, a combination of TsFontStyle elements
  @param AColor     RGB valoe of the font color
  @param APosition  Specifies subscript or superscript text.
  @returns          Index of the font in the workbook's font list
-------------------------------------------------------------------------------}
function TsWorkbook.AddFont(const AFontName: String; ASize: Single;
  AStyle: TsFontStyles; AColor: TsColor;
  APosition: TsFontPosition = fpNormal): Integer;
var
  fnt: TsFont;
begin
  fnt := TsFont.Create;
  fnt.FontName := AFontName;
  fnt.Size := ASize;
  fnt.Style := AStyle;
  fnt.Color := AColor;
  fnt.Position := APosition;
  Result := AddFont(fnt);
end;


{@@ ----------------------------------------------------------------------------
  Adds a font to the font list. Returns the index in the font list.

  @param AFont      TsFont record containing all font parameters
  @returns          Index of the font in the workbook's font list
-------------------------------------------------------------------------------}
function TsWorkbook.AddFont(const AFont: TsFont): Integer;
begin
  result := FFontList.Add(AFont);
end;


{@@ ----------------------------------------------------------------------------
  Creates a new font as a copy of the font at the specified index.
  The new font is NOT YET added to the font list.
  If the user does not add the font to the font list he is responsibile for
  destroying it.
-------------------------------------------------------------------------------}
function TsWorkbook.CloneFont(const AFontIndex: Integer): TsFont;
var
  fnt: TsFont;
begin
  Result := TsFont.Create;
  fnt := GetFont(AFontIndex);
  Result.FontName := fnt.FontName;
  Result.Size := fnt.Size;
  Result.Style := fnt.Style;
  Result.Color := fnt.Color;
  Result.Position := fnt.Position;
end;


{@@ ----------------------------------------------------------------------------
  Deletes a font.
  Use with caution because this will screw up the font assignment to cells.
  The only legal reason to call this method is from a reader of a file format
  in which the missing font #4 of BIFF does exist.
-------------------------------------------------------------------------------}
procedure TsWorkbook.DeleteFont(const AFontIndex: Integer);
var
  fnt: TsFont;
begin
  if AFontIndex < FFontList.Count then
  begin
    fnt := TsFont(FFontList.Items[AFontIndex]);
    if fnt <> nil then fnt.Free;
    FFontList.Delete(AFontIndex);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Checks whether the font with the given specification is already contained in
  the font list. Returns the index, or -1 if not found.

  @param AFontName  Name of the font (like 'Arial')
  @param ASize      Size of the font in points
  @param AStyle     Style of the font, a combination of TsFontStyle elements
  @param AColor     RGB value of the font color
  @param APosition  Specified subscript or superscript text.
  @returns          Index of the font in the font list, or -1 if not found.
-------------------------------------------------------------------------------}
function TsWorkbook.FindFont(const AFontName: String; ASize: Single;
  AStyle: TsFontStyles; AColor: TsColor; APosition: TsFontPosition = fpNormal): Integer;
begin
  Result := FindFontInList(FFontList, AFontName, ASize, AStyle, AColor, APosition);
end;


{@@ ----------------------------------------------------------------------------
  Returns the count of built-in fonts (default font, hyperlink font, bold font
  by default).
-------------------------------------------------------------------------------}
function TsWorkbook.GetBuiltinFontCount: Integer;
begin
  Result := FBuiltinFontCount;
end;


{@@ ----------------------------------------------------------------------------
  Returns the default font. This is the first font (index 0) in the font list
-------------------------------------------------------------------------------}
function TsWorkbook.GetDefaultFont: TsFont;
begin
  Result := GetFont(0);
end;


{@@ ----------------------------------------------------------------------------
  Returns the point size of the default font
-------------------------------------------------------------------------------}
function TsWorkbook.GetDefaultFontSize: Single;
begin
  Result := GetFont(0).Size;
end;


{@@ ----------------------------------------------------------------------------
  Returns the font with the given index.

  @param   AIndex   Index of the font to be considered
  @returns @link(TsFont) instance containing all parameters of the font (or nil if not found).
-------------------------------------------------------------------------------}
function TsWorkbook.GetFont(AIndex: Integer): TsFont;
begin
  if (AIndex >= 0) and (AIndex < FFontList.Count) then
    Result := FFontList.Items[AIndex]
  else
    Result := nil;
end;


{@@ ----------------------------------------------------------------------------
  Returns a string which identifies the font with a given index.

  @param   AIndex    Index of the font
  @returns String with font name, font size etc.
-------------------------------------------------------------------------------}
function TsWorkbook.GetFontAsString(AIndex: Integer): String;
begin
  Result := fpsUtils.GetFontAsString(GetFont(AIndex));
end;


{@@ ----------------------------------------------------------------------------
  Returns the count of registered fonts
-------------------------------------------------------------------------------}
function TsWorkbook.GetFontCount: Integer;
begin
  Result := FFontList.Count;
end;


{@@ ----------------------------------------------------------------------------
  Initializes the font list by adding 5 fonts:

    0: default font
    1: like default font, but blue and underlined (for hyperlinks)
    2: like default font, but bold
    3: like default font, but italic
-------------------------------------------------------------------------------}
procedure TsWorkbook.InitFonts;
var
  fntName: String;
  fntSize: Single;
begin
  // Memorize old default font
  with TsFont(FFontList.Items[0]) do
  begin
    fntName := FontName;
    fntSize := Size;
  end;

  // Remove current font list
  RemoveAllFonts;

  // Build new font list
  SetDefaultFont(fntName, fntSize);                      // FONT0: Default font
  AddFont(fntName, fntSize, [fssUnderline], scBlue);     // FONT1: Hyperlink font = blue & underlined
  AddFont(fntName, fntSize, [fssBold], scBlack);         // FONT2: Bold font
  AddFont(fntName, fntSize, [fssItalic], scBlack);       // FONT3: Italic font (not used directly)

  FBuiltinFontCount := FFontList.Count;
end;


{@@ ----------------------------------------------------------------------------
  Clears the list of fonts and releases their memory.
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveAllFonts;
var
  i: Integer;
  fnt: TsFont;
begin
  for i := FFontList.Count-1 downto 0 do
  begin
    fnt := TsFont(FFontList.Items[i]);
    fnt.Free;
    FFontList.Delete(i);
  end;
  FBuiltinFontCount := 0;
end;


{@@ ----------------------------------------------------------------------------
  Replaces the built-in font at a specific index with different font parameters
-------------------------------------------------------------------------------}
procedure TsWorkbook.ReplaceFont(AFontIndex: Integer; AFontName: String;
  ASize: Single; AStyle: TsFontStyles; AColor: TsColor;
  APosition: TsFontPosition = fpNormal);
var
  fnt: TsFont;
begin
  if (AFontIndex < FBuiltinFontCount) then //and (AFontIndex <> 4) then
  begin
    fnt := TsFont(FFontList[AFontIndex]);
    fnt.FontName := AFontName;
    fnt.Size := ASize;
    fnt.Style := AStyle;
    fnt.Color := AColor;
    fnt.Position := APosition;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Defines the default font. This is the font with index 0 in the FontList.
  The next built-in fonts will have the same font name and size
-------------------------------------------------------------------------------}
procedure TsWorkbook.SetDefaultFont(const AFontName: String; ASize: Single);
var
  i: Integer;
begin
  if FFontList.Count = 0 then
    AddFont(AFontName, ASize, [], scBlack)
  else
  for i:=0 to FBuiltinFontCount-1 do
    if (i <> 4) and (i < FFontList.Count) then           // wp: why if font #4 relevant here ????
      with TsFont(FFontList[i]) do
      begin
        FontName := AFontName;
        Size := ASize;
      end;
end;


