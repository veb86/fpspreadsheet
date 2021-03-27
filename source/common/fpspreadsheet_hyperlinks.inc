{ Included by fpspreadsheet.pas }

{ Contains code for hyperlinks }

{@@ ----------------------------------------------------------------------------
  Checks whether the specified cell contains a hyperlink and returns a pointer
  to the hyperlink data.

  @param   ACell  Pointer to the cell
  @returns Pointer to the @link(TsHyperlink) record, or NIL if the cell does not contain a hyperlink.
-------------------------------------------------------------------------------}
function TsWorksheet.FindHyperlink(ACell: PCell): PsHyperlink;
begin
  if HasHyperlink(ACell) then
    Result := PsHyperlink(FHyperlinks.FindByRowCol(ACell^.Row, ACell^.Col))
  else
    Result := nil;
end;


{@@ ----------------------------------------------------------------------------
  Reads the hyperlink information of a specified cell.

  @param   ACell         Pointer to the cell considered
  @returns @link(TsHyperlink) record with the hyperlink data assigned to the cell. If the cell is not a hyperlink the result field Kind is hkNone.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadHyperlink(ACell: PCell): TsHyperlink;
var
  hyperlink: PsHyperlink;
begin
  hyperlink := FindHyperlink(ACell);
  if hyperlink <> nil then
    Result := hyperlink^
  else
  begin
    Result.Row := ACell^.Row;
    Result.Col := ACell^.Col;
    Result.Target := '';
    Result.Tooltip := '';
  end;
end;


{@@ ----------------------------------------------------------------------------
  Removes a hyperlink from the specified cell. Releaes memory occupied by
  the associated TsHyperlink record. Cell content type is converted to
  cctUTF8String.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveHyperlink(ACell: PCell);
begin
  if HasHyperlink(ACell) then
  begin
    FHyperlinks.DeleteHyperlink(ACell^.Row, ACell^.Col);
    Exclude(ACell^.Flags, cfHyperlink);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Checks whether the passed string represents a valid hyperlink target

  The string must either be
  * a fully qualified URI, 
  * a local relative (!) file name, or 
  * a # followed by a cell address in the current workbook
  
  @param   AValue  String to be checked. 
  @param   AErrMsg Error message in case that the string is not correct.
  @returns @TRUE if the string is correct, @FALSE otherwise
-------------------------------------------------------------------------------}
function TsWorksheet.ValidHyperlink(AValue: String; out AErrMsg: String): Boolean;
var
  u: TUri;
  sheet: TsWorksheet;
  r, c: Cardinal;
begin
  Result := false;
  AErrMsg := '';
  if AValue = '' then
  begin
    AErrMsg := rsEmptyHyperlink;
    exit;
  end else
  if (AValue[1] = '#') then
  begin
    Delete(AValue, 1, 1);
    if not FWorkbook.TryStrToCell(AValue, sheet, r, c) then
    begin
      AErrMsg := Format(rsNoValidHyperlinkInternal, ['#'+AValue]);
      exit;
    end;
  end else
  begin
    u := ParseURI(AValue);
    if SameText(u.Protocol, 'mailto') then
    begin
      Result := true;   // To do: Check email address here...
      exit;
    end else
    if SameText(u.Protocol, 'file') then
    begin
      if FilenameIsAbsolute(u.Path + u.Document) then
      begin
        Result := true;
        exit;
      end else
      begin
        AErrMsg := Format(rsLocalfileHyperlinkAbs, [AValue]);
        exit;
      end;
    end else
    begin
      Result := true;
      exit;
    end;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Assigns a hyperlink to the cell at the specified row and column
  Cell content is not affected by the presence of a hyperlink.

  @param  ARow          Row index of the cell considered
  @param  ACol          Column index of the cell considered
  @param  ATarget       Hyperlink address given as a fully qualitifed URI for external links, or as a # followed by a cell address for internal links.
  @param  ATooltip      Text for popup tooltip hint used by Excel
  @returns Pointer to the cell with the hyperlink
-------------------------------------------------------------------------------}
function TsWorksheet.WriteHyperlink(ARow, ACol: Cardinal; ATarget: String;
  ATooltip: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteHyperlink(Result, ATarget, ATooltip);
end;


{@@ ----------------------------------------------------------------------------
  Assigns a hyperlink to the specified cell.

  @param  ACell         Pointer to the cell considered
  @param  ATarget       Hyperlink address given as a fully qualitifed URI for external links, or as a # followed by a cell address for internal links. Local files can be specified also by their name relative to the workbook. An existing hyperlink is removed if ATarget is empty.
  @param  ATooltip      Text for popup tooltip hint used by Excel
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteHyperlink(ACell: PCell; ATarget: String;
  ATooltip: String = '');

  function GetDisplayText(ATarget: String): String;
  var
    target, bm: String;
  begin
    SplitHyperlink(ATarget, target, bm);
    if pos('file:', lowercase(ATarget))=1 then
    begin
      URIToFilename(target, Result);
      ForcePathDelims(Result);
      if bm <> '' then Result := Result + '#' + bm;
    end else
    if target = '' then
      Result := bm
    else
      Result := ATarget;
  end;

var
  fmt: TsCellFormat;
  noCellText: Boolean = false;
begin
  if ACell = nil then
    exit;

  fmt := ReadCellFormat(ACell);

  // Empty target string removes the hyperlink. Resets the font from hyperlink
  // to default font.
  if ATarget = '' then begin
    RemoveHyperlink(ACell);
    if fmt.FontIndex = HYPERLINK_FONTINDEX then
      WriteFont(ACell, DEFAULT_FONTINDEX);
    exit;
  end;

  // Detect whether the cell already has a hyperlink, but has no other content.
  if HasHyperlink(ACell) then
    noCellText := (ACell^.ContentType = cctUTF8String) and
      (GetDisplayText(ReadHyperlink(ACell).Target) = ReadAsText(ACell));

  // Attach the hyperlink to the cell
  FHyperlinks.AddHyperlink(ACell^.Row, ACell^.Col, ATarget, ATooltip);
  Include(ACell^.Flags, cfHyperlink);

  // If there is no other cell content use the target as cell label string.
  if (ACell^.ContentType = cctEmpty) or noCellText then
  begin
    ACell^.ContentType := cctUTF8String;
    ACell^.UTF8StringValue := GetDisplayText(ATarget);
  end;

  // Select the hyperlink font.
  if fmt.FontIndex = DEFAULT_FONTINDEX then
  begin
    fmt.FontIndex := HYPERLINK_FONTINDEX;
    Include(fmt.UsedFormattingFields, uffFont);
    ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt);
  end;

  ChangedCell(ACell^.Row, ACell^.Col);
end;


{==============================================================================}
{  TsWorkbook code for hyperlinls                                              }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Returns the hyperlink font. This is the font with index 6 in the font list
-------------------------------------------------------------------------------}
function TsWorkbook.GetHyperlinkFont: TsFont;
begin
  Result := GetFont(HYPERLINK_FONTINDEX);
end;


