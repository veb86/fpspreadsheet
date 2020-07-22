{ Included by fpspreadsheet.pas }

{ Clipboard access }

{@@ ----------------------------------------------------------------------------
  Writes the selected cells to a stream for usage in the clipboard.
  Transfer to the clipboard has do be done by the calling routine since
  fpspreadsheet does not "know" the system's clipboard.
-------------------------------------------------------------------------------}
procedure TsWorkbook.CopyToClipboardStream(AStream: TStream;
  AFormat: TsSpreadsheetFormat; AParams: TsStreamParams = []);
var
  clipbook: TsWorkbook;
  clipsheet: TsWorksheet;
  sel: TsCellRange;
  range: TsCellRangeArray;
  r, c: Cardinal;
  srccell, destcell: PCell;
begin
  if AStream = nil then
    exit;

  if ActiveWorksheet = nil then
    exit;

  // Create workbook which will be written to clipboard stream
  // Contains only the selected worksheet and the selected cells.
  clipbook := TsWorkbook.Create;
  try
    clipsheet := clipbook.AddWorksheet(ActiveWorksheet.Name);
    for sel in ActiveWorksheet.GetSelection do
    begin
      for r := sel.Row1 to sel.Row2 do
        for c := sel.Col1 to sel.Col2 do
        begin
          srccell := ActiveWorksheet.FindCell(r, c);
          if ActiveWorksheet.IsMerged(srccell) then
            srccell := ActiveWorksheet.FindMergeBase(srccell);
          if srccell <> nil then begin
            destcell := clipsheet.GetCell(r, c);  // wp: why was there AddCell?
            clipsheet.CopyCell(srccell, destcell);
          end;
        end;
    end;
    // Select the same cells as in the source workbook.
    range := ActiveWorksheet.GetSelection;
    clipsheet.SetSelection(range);
    clipsheet.SelectCell(range[0].Row1, range[0].Col1);

    // Write this workbook to a stream. Set the parameter spClipboard to
    // indicate that this should be the special clipboard version of the stream.
    clipbook.WriteToStream(AStream, AFormat, AParams + [spClipboard]);

    if AFormat = sfCSV then
      AStream.WriteByte(0);

    // The calling routine which copies the stream to the clipboard requires
    // the stream to be at its beginning.
    AStream.Position := 0;
  finally
    clipbook.Free;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Copies the cells stored in the specified stream to the active worksheet.
  The provided stream contains data from the system's clipboard.
  Note that transfer from the clipboard to the stream has to be done by the
  calling routine since fpspreadsheet does not "know" the system's clipboard.
-------------------------------------------------------------------------------}
procedure TsWorkbook.PasteFromClipboardStream(AStream: TStream;
  AFormat: TsSpreadsheetFormat; AOperation: TsCopyOperation;
  AParams: TsStreamParams = []; ATransposed: Boolean = false);
var
  clipbook: TsWorkbook;
  clipsheet: TsWorksheet;
  sel: TsCellRange;
  selArray: TsCellRangeArray = nil;
  r, c: LongInt;
  dr, dc: LongInt;
  srcCell, destCell: PCell;
  i: Integer;              // counter
  ncs, nrs: Integer;       // Num cols source, num rows source, ...
  //ncd, nrd: Integer;
  rdest, cdest: Integer;   // row and column index at destination
  nselS, nselD: Integer;   // count of selected blocks
begin
  Unused(ATransposed);

  if AStream = nil then
    exit;

  if ActiveWorksheet = nil then
    exit;

  if AOperation = coNone then
    exit;

  // Create workbook into which the clipboard stream will write
  clipbook := TsWorkbook.Create;
  try
    clipbook.Options := clipbook.Options + [boReadFormulas];
    // Read stream into this temporary workbook
    // Set last parameter (ClipboardMode) to TRUE to activate special format
    // treatment for clipboard, if needed.
    clipbook.ReadFromStream(AStream, AFormat, AParams + [spClipboard]);
    clipsheet := clipbook.GetWorksheetByIndex(0);

    // count of blocks in source (clipboard sheet)
    nselS := clipsheet.GetSelectionCount;
    // count of selected blocks at destination
    nselD := ActiveWorksheet.GetSelectionCount;

    // -------------------------------------------------------------------------
    // Case (1): Destination is a single cell, source can be any shape
    //           --> Source shape is duplicated starting at destination
    // -------------------------------------------------------------------------
    if (nselD = 1)
       and (ActiveWorksheet.GetSelection[0].Col1 = ActiveWorksheet.GetSelection[0].Col2)
       and (ActiveWorksheet.GetSelection[0].Row1 = ActiveWorksheet.GetSelection[0].Row2)
    then begin
      // Find offset of active cell to left/top cell in clipboard sheet
      dr := LongInt(ActiveWorksheet.ActiveCellRow) - clipsheet.ActiveCellRow;
      dc := LongInt(ActiveWorksheet.ActiveCellCol) - clipsheet.ActiveCellCol;
      // Copy cells from clipboard sheet to active worksheet
      // Shift them such that top/left of clipboard sheet is at active cell
      for srcCell in clipsheet.Cells do
      begin
        r := LongInt(srcCell^.Row) + dr;
        c := LongInt(srcCell^.Col) + dc;
        destcell := ActiveWorksheet.GetCell(r, c);
        case AOperation of
          coCopyCell    : ActiveWorksheet.CopyCell(srcCell, destCell);
          coCopyValue   : ActiveWorksheet.CopyValue(srcCell, destCell);
          coCopyFormat  : ActiveWorksheet.CopyFormat(srcCell, destCell);
          coCopyFormula : ActiveWorksheet.CopyFormula(srcCell, destCell);
        end;
      end;
      // Select all copied cells
      sel := Range(Cardinal(-1), Cardinal(-1), Cardinal(-1), Cardinal(-1));
      SetLength(selArray, nselS);
      for i := 0 to nselS-1 do
      begin
        sel := clipsheet.GetSelection[i];
        selArray[i].Row1 := LongInt(sel.Row1) + dr;
        selArray[i].Col1 := LongInt(sel.Col1) + dc;
        selArray[i].Row2 := LongInt(sel.Row2) + dr;
        selArray[i].Col2 := LongInt(sel.Col2) + dc;
      end;
      ActiveWorksheet.SetSelection(selArray);
      // Select active cell. If not found in the file, let's use the last cell of the selections
      if (clipsheet.ActiveCellRow <> 0) and (clipsheet.ActiveCellCol <> 0) then
      begin
        r := clipsheet.ActiveCellRow;
        c := clipsheet.ActiveCellCol;
      end else
      begin
        r := LongInt(sel.Row2);
        c := LongInt(sel.Col2);
      end;
      if (r <> -1) and (c <> -1) then
        ActiveWorksheet.SelectCell(r + dr, c + dc);
    end
    else
    // -------------------------------------------------------------------------
    // Case (2): Source is a single block (not necessarily a cell), Dest can be
    //           any shape --> source is tiled into destination
    // -------------------------------------------------------------------------
//    if nselS = 1 then
    begin
      // size of source block
      with clipsheet do
      begin
        ncs := LongInt(GetLastColIndex(true)) - LongInt(GetFirstColIndex(true)) + 1;
        nrs := LongInt(GetLastRowIndex(true)) - LongInt(GetFirstRowIndex(true)) + 1;
      end;
      // Iterate over all destination blocks
      for i := 0 to nselD-1 do
      begin
        r := ActiveWorksheet.GetSelection[i].Row1;
        while r <= longint(ActiveWorksheet.GetSelection[i].Row2) do begin
          c := ActiveWorksheet.GetSelection[i].Col1;
          while c <= longint(ActiveWorksheet.GetSelection[i].Col2) do begin
            dr := r - clipsheet.GetFirstRowIndex;
            dc := c - clipsheet.GetFirstColIndex;
            for srccell in clipsheet.Cells do
            begin
              rdest := longint(srccell^.Row) + dr;
              if rdest > integer(ActiveWorksheet.GetSelection[i].Row2) then
                Continue;
              cdest := longint(srcCell^.Col) + dc;
              if cdest > integer(ActiveWorksheet.GetSelection[i].Col2) then
                Continue;
              destcell := ActiveWorksheet.GetCell(
                LongInt(srcCell^.Row) + dr,
                LongInt(srcCell^.Col) + dc
              );
              case AOperation of
                coCopyCell    : ActiveWorksheet.CopyCell(srcCell, destCell);
                coCopyValue   : ActiveWorksheet.CopyValue(srcCell, destCell);
                coCopyFormat  : ActiveWorksheet.CopyFormat(srcCell, destCell);
                coCopyFormula : ActiveWorksheet.CopyFormula(srcCell, destCell);
              end;
            end;  // for srcCell
            inc(c, ncs);
          end;  // while c...
          inc(r, nrs);
        end;  // while r...
      end;  // for i
      // No need to select copied cells - they already are.
    end ;
  finally
    clipbook.Free;
  end;
end;


