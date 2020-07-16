{ Included by fpspreadsheet.pas }

{ Contains code for comments }


{@@ ----------------------------------------------------------------------------
  Checks whether a cell contains a comment and returns a pointer to the
  comment data.

  @param  ACell  Pointer to the cell
  @return Pointer to the TsComment record (nil, if the cell does not have a
          comment)
-------------------------------------------------------------------------------}
function TsWorksheet.FindComment(ACell: PCell): PsComment;
begin
  if HasComment(ACell) then
    Result := PsComment(FComments.FindByRowCol(ACell^.Row, ACell^.Col))
  else
    Result := nil;
end;


{@@ ----------------------------------------------------------------------------
  Checks whether a specific cell contains a comment
-------------------------------------------------------------------------------}
function TsWorksheet.HasComment(ACell: PCell): Boolean;
begin
  Result := (ACell <> nil) and (cfHasComment in ACell^.Flags);
end;


{@@ ----------------------------------------------------------------------------
  Returns the comment text attached to a specific cell

  @param  ARow   (0-based) index to the row
  @param  ACol   (0-based) index to the column
  @return Text assigned to the cell as a comment
-------------------------------------------------------------------------------}
function TsWorksheet.ReadComment(ARow, ACol: Cardinal): String;
var
  comment: PsComment;
begin
  Result := '';
  comment := PsComment(FComments.FindByRowCol(ARow, ACol));
  if comment <> nil then
    Result := comment^.Text;
end;


{@@ ----------------------------------------------------------------------------
  Returns the comment text attached to a specific cell

  @param  ACell  Pointer to the cell
  @return Text assigned to the cell as a comment
-------------------------------------------------------------------------------}
function TsWorksheet.ReadComment(ACell: PCell): String;
var
  comment: PsComment;
begin
  Result := '';
  comment := FindComment(ACell);
  if comment <> nil then
    Result := comment^.Text;
end;


{@@ ----------------------------------------------------------------------------
  Removes the comment from a cell and releases the memory occupied by the node.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveComment(ACell: PCell);
begin
  if HasComment(ACell) then
  begin
    FComments.DeleteComment(ACell^.Row, ACell^.Col);
    Exclude(ACell^.Flags, cfHasComment);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Adds a comment to a specific cell

  @param  ARow   (0-based) row index of the cell
  @param  ACol   (0-based) column index of the cell
  @param  AText  Comment text
  @return Pointer to the cell containing the comment
-------------------------------------------------------------------------------}
function TsWorksheet.WriteComment(ARow, ACol: Cardinal; AText: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteComment(Result, AText);
end;


{@@ ----------------------------------------------------------------------------
  Adds a comment to a specific cell

  @param  ACell  Pointer to the cell
  @param  AText  Comment text
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteComment(ACell: PCell; AText: String);
begin
  if ACell = nil then
    exit;

  // Remove the comment if an empty string is passed
  if AText = '' then
  begin
    RemoveComment(ACell);
    exit;
  end;

  // Add new comment record
  FComments.AddComment(ACell^.Row, ACell^.Col, AText);
  Include(ACell^.Flags, cfHasComment);

  ChangedCell(ACell^.Row, ACell^.Col);

end;

