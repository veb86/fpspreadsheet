{ Included by fpspreadsheet.pas }

{ Code for embedded objects (images) }

{==============================================================================}
{   TsWorksheet code for embedded objects                                      }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Calculates the position of the image with given index relative to the cell
  containing the top/left corner of the image.

  @param  x    worksheet-relative coordinate of the left image edge, in workbook units
  @param  y    worksheet-relative coordinate of the top image edge, in workbook units
  @param  ARow Index of the row containing the top/left corner of the image
  @param  ACol Index of the column containing the top/left corner of the image
  @param  ARowOffset  Distance, in workbook units, between top cell and image borders
  @param  AColOffset  Distance, in workbook units, between left cell and image borders
  @param  AScaleX  Scaling factor for the image width
  @param  AScaleY  Scaling factor for the image height
-------------------------------------------------------------------------------}
procedure TsWorksheet.CalcImageCell(AIndex: Integer; x, y, AWidth, AHeight: Double;
  out ARow, ACol: Cardinal; out ARowOffs, AColOffs, AScaleX, AScaleY: Double);
// All lengths are in workbook units!
var
  colW, rowH, sum: Double;
  embobj: TsEmbeddedObj;
begin
  ACol := 0;
  sum := 0;
  colW := GetColWidth(0, FWorkbook.Units);
  while (sum + colW < x) do begin
    sum := sum + colW;
    inc(ACol);
    colW := GetColWidth(ACol, FWorkbook.Units);
  end;
  AColOffs := x - sum;

  ARow := 0;
  sum := 0;
  rowH := CalcRowHeight(0);
  while (sum + rowH < y) do begin
    sum := sum + rowH;
    inc(ARow);
    rowH := CalcRowHeight(ARow);
  end;
  ARowOffs := y - sum;

  embObj := FWorkbook.GetEmbeddedObj(AIndex);
  AScaleX := AWidth / embObj.ImageWidth;
  AScaleY := AHeight / embObj.ImageHeight;
end;


{@@ ----------------------------------------------------------------------------
  Calculates image extent

  @param  AIndex    Index of the image into the worksheet's image list
  @param  UsePixels If @TRUE then pixels are used for calculation - this improves the display of the images in Excel
  @param  ARow1     Index of the row containing the top edge of the image
  @param  ACol1     Index of the column containing the left edege of the image
  @param  ARow2     Index of the row containing the right edge of the image
  @param  ACol2     Index of the column containing the bottom edge of the image
  @param  ARowOffs1 Distance between the top edge of image and row 1
  @param  AColOffs1 Distance between the left edge of image and column 1
  @param  ARowOffs2 Distance between the bottom edge of image and top of row 2
  @param  AColOffs2 Distance between the right edge of image and left of col 2
  @param  x         Absolute coordinate of left edge of image
  @param  y         Absolute coordinate of top edge of image
  @param  AWidth    Width of the image
  @param  AHeight   Height of the image

  All dimensions are in workbook units
-------------------------------------------------------------------------------}
procedure TsWorksheet.CalcImageExtent(AIndex: Integer; UsePixels: Boolean;
  out ARow1, ACol1, ARow2, ACol2: Cardinal;
  out ARowOffs1, AColOffs1, ARowOffs2, AColOffs2: Double;
  out x,y, AWidth, AHeight: Double);
var
  img: TsImage;
  obj: TsEmbeddedObj;
  colW, rowH: Double;
  totH: Double;
  r, c: Integer;
  w_px, h_px: Integer;
  totH_px, rowH_px: Integer;
  totW_px, colW_px: Integer;
  ppi: Integer;
  u: TsSizeUnits;
begin
  // Abbreviations
  ppi := ScreenPixelsPerInch;
  u := FWorkbook.Units;

  img := GetImage(AIndex);
  ARow1 := img.Row;
  ACol1 := img.Col;
  ARowOffs1 := img.OffsetX;                 // in workbook units
  AColOffs1 := img.OffsetY;                 // in workbook units

  obj := FWorkbook.GetEmbeddedObj(img.Index);
  AWidth := obj.ImageWidth * img.ScaleX;    // in workbook units
  AHeight := obj.ImageHeight * img.ScaleY;  // in workbook units

  // Find x coordinate of left image edge, in workbook units
  x := AColOffs1;
  for c := 0 to ACol1-1 do
  begin
    colW := GetColWidth(c, u);
    x := x + colW;
  end;
  // Find y coordinate of top image edge, in workbook units.
  y := ARowOffs1;
  for r := 0 to ARow1 - 1 do
  begin
    rowH := CalcRowHeight(r);
    y := y + rowH;
  end;

  if UsePixels then
  // Use pixels for calculation. Better for Excel, maybe due to rounding error?
  begin
    // If we don't know the ppi of the screen the calculation is not exact!
    w_px := ptsToPx(FWorkbook.ConvertUnits(AWidth, u, suPoints), ppi);
    h_px := ptsToPx(FWorkbook.ConvertUnits(AHeight, u, suPoints), ppi);
    // Find cell with right image edge. Find horizontal within-cell-offsets
    totW_px := -ptsToPx(FWorkbook.ConvertUnits(AColOffs1, u, suPoints), ppi);
    ACol2 := ACol1;
    while (totW_px < w_px) do
    begin
      colW := GetColWidth(ACol2, u);
      colW_px := ptsToPx(FWorkbook.ConvertUnits(colW, u, suPoints), ppi);
      totW_px := totW_px + colW_px;
      if totW_px > w_px then
      begin
        AColOffs2 := FWorkbook.ConvertUnits(pxToPts(colW_px - (totW_px - w_px), ppi), suPoints, u);
        break;
      end;
      inc(ACol2);
    end;
    // Find cell with bottom image edge. Find vertical within-cell-offset.
    totH_px := -ptsToPx(FWorkbook.ConvertUnits(ARowOffs1, u, suPoints), ppi);
    ARow2 := ARow1;
    while (totH_px < h_px) do
    begin
      rowH := CalcRowHeight(ARow2);
      rowH_px := ptsToPx(FWorkbook.ConvertUnits(rowH, u, suPoints), ppi);
      totH_px := totH_px + rowH_px;
      if totH_px > h_px then
      begin
        ARowOffs2 := FWorkbook.ConvertUnits(pxToPts(rowH_px - (totH_px - h_px), ppi), suPoints, u);
        break;
      end;
      inc(ARow2);
    end;
  end
  else    // Use workbook units for calculation
  begin
    // Find cell with right image edge. Find horizontal within-cell-offsets
    totH := -ARowOffs1;
    ARow2 := ARow1;
    while (totH < AHeight) do
    begin
      rowH := CalcRowHeight(ARow2);
      totH := totH + rowH;
      if totH >= AHeight then
      begin
        ARowOffs2 := rowH - (totH - AHeight);
        break;
      end;
      inc(ARow2);
    end;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Returns the parameters of the image stored in the internal image list at
  the specified index.

  @param    AIndex      Index of the image to be retrieved
  @returns  @link(TsImage) record with all image parameters.
-------------------------------------------------------------------------------}
function TsWorksheet.GetImage(AIndex: Integer): TsImage;
var
  img: PsImage;
begin
  img := PsImage(FImages[AIndex]);
  Result := img^;
end;


{@@ ----------------------------------------------------------------------------
  Returns the count of images that are embedded into this sheet.
-------------------------------------------------------------------------------}
function TsWorksheet.GetImageCount: Integer;
begin
  Result := FImages.Count;
end;


function TsWorksheet.GetPointerToImage(AIndex: Integer): PsImage;
begin
  Result := PsImage(FImages[AIndex]);
end;


{@@ ----------------------------------------------------------------------------
  Removes all image from the internal image list.
  The image streams (stored by the workbook), however, are retained because
  images may also be used as header/footer images.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveAllImages;
var
  i: Integer;
begin
  for i := FImages.Count-1 downto 0 do
    RemoveImage(i);
end;


{@@ ----------------------------------------------------------------------------
  Removes an image from the internal image list.
  The image is identified by its index.
  The image stream (stored by the workbook) is retained.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveImage(AIndex: Integer);
var
  img: PsImage;
begin
  img := PsImage(FImages[AIndex]);
  if (img <> nil) then begin
    if (img^.Picture <> nil) then img^.Picture.Free;
    img^.HyperlinkTarget := '';
    img^.HyperlinkToolTip := '';
  end;
  Dispose(img);
  FImages.Delete(AIndex);
end;


{@@ ----------------------------------------------------------------------------
  Adds an embedded image to the worksheet

  @param  ARow       Index of the row at which the image begins (top edge)
  @param  ACol       Index of the column at which the image begins (left edge)
  @param  AFileName  Name of the image file
  @param  AOffsetX   The image is offset horizontally from the left edge of the anchor cell. May reach into another cell. Value is in workbook units.
  @param  AOffsetY   The image is offset vertically from the top edge of the anchor cell. May reach into another cell. Value is in workbook units.
  @param  AScaleX    Horizontal scaling factor of the image
  @param  AScaleY    Vertical scaling factor of the image
  @returns Index into the internal image list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteImage(ARow, ACol: Cardinal; AFileName: String;
  AOffsetX: Double = 0.0; AOffsetY: Double = 0.0;
  AScaleX: Double = 1.0; AScaleY: Double = 1.0): Integer;
var
  idx: Integer;
begin
  // Does the image already exist?
  idx := Workbook.FindEmbeddedObj(AFileName);
  // No? Open and store in embedded object list.
  if idx = -1 then
    idx := Workbook.AddEmbeddedObj(AFileName);
  // An error has occured? Error is already logged. Just exit.
  if idx = -1 then
    exit;

  // Everything ok here...
  Result := WriteImage(ARow, ACol, idx, AOffsetX, AOffsetY, AScaleX, AScaleY);
end;


{@@ ----------------------------------------------------------------------------
  Adds an embedded image to the worksheet. The image passed in a stream.

  @param   ARow       Index of the row at which the image begins (top edge)
  @param   ACol       Index of the column at which the image begins (left edge)
  @param   AStream    Stream which contains the image data
  @param   AOffsetX   The image is offset horizontally from the left edge of the anchor cell. May reach into another cell. Value is in workbook units.
  @param   AOffsetY   The image is offset vertically from the top edge of the anchor cell. May reach into another cell. Value is in workbook units.
  @param   AScaleX    Horizontal scaling factor of the image
  @param   AScaleY    Vertical scaling factor of the image
  @param   ASize      Number ob bytes to be read from the input stream.
  @returns Index into the internal image list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteImage(ARow, ACol: Cardinal; AStream: TStream;
  AOffsetX: Double = 0.0; AOffsetY: Double = 0.0;
  AScaleX: Double = 1.0; AScaleY: Double = 1.0;
  ASize: Int64 = -1): Integer;
var
  idx: Integer;
begin
  // Copy the stream to a new item in embedded object list.
  idx := Workbook.AddEmbeddedObj(AStream, '', ASize);

  // An error has occured? Error is already logged. Just exit.
  if idx = -1 then
    exit;

  // Everything ok here...
  Result := WriteImage(ARow, ACol, idx, AOffsetX, AOffsetY, AScaleX, AScaleY);
end;


function TsWorksheet.WriteImage(ARow, ACol: Cardinal; AImageIndex: Integer;
  AOffsetX: Double = 0.0; AOffsetY: Double = 0.0;
  AScaleX: Double = 1.0; AScaleY: Double = 1.0): Integer;
var
  img: PsImage;
begin
  New(img);
  InitImageRecord(img^, ARow, ACol, AOffsetX, AOffsetY, AScaleX, AScaleY);
  img^.Index := AImageIndex;
  Result := FImages.Add(img);
end;



{==============================================================================}
{   TsWorkbook code for embedded objects                                       }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Creates a new "embedded" stream and loads the specified file.
  Returns the index of the embedded file item.
  Image dimensions are converted to workbook units.
-------------------------------------------------------------------------------}
function TsWorkbook.AddEmbeddedObj(const AFileName: String): Integer;
var
  obj: TsEmbeddedObj = nil;
begin
  if not FileExists(AFileName) then
  begin
    AddErrorMsg(rsFileNotFound, [AFileName]);
    Result := -1;
    exit;
  end;

  obj := TsEmbeddedObj.Create;
  if obj.LoadFromFile(AFileName) then
  begin
    obj.ImageWidth := ConvertUnits(obj.ImageWidth, suInches, FUnits);
    obj.ImageHeight := ConvertUnits(obj.ImageHeight, suInches, FUnits);
    Result := FEmbeddedObjList.Add(obj)
  end else
  begin
    AddErrorMsg(rsFileFormatNotSupported, [AFileName]);
    obj.Free;
    Result := -1;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Creates a new "embedded" stream and copies the specified stream to it.
  Returns the index of the embedded object.
-------------------------------------------------------------------------------}
function TsWorkbook.AddEmbeddedObj(AStream: TStream;
  const AName: String = ''; ASize: Int64 = -1): Integer;
var
  obj: TsEmbeddedObj = nil;
begin
  obj := TsEmbeddedObj.Create;
  if obj.LoadFromStream(AStream, AName, ASize) then
  begin
    obj.ImageWidth := ConvertUnits(obj.ImageWidth, suInches, FUnits);
    obj.ImageHeight := ConvertUnits(obj.ImageHeight, suInches, FUnits);
    Result := FEmbeddedObjList.Add(obj)
  end else
  begin
    AddErrorMsg(rsImageFormatNotSupported);
    obj.Free;
    Result := -1;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Checks whether an embedded object with the specified file name already exists.
  If yes, returns its index in the object list, or -1 if no.
-------------------------------------------------------------------------------}
function TsWorkbook.FindEmbeddedObj(const AFileName: String): Integer;
var
  obj: TsEmbeddedObj;
begin
  for Result:=0 to FEmbeddedObjList.Count-1 do
  begin
    obj := TsEmbeddedObj(FEmbeddedObjList[Result]);
    if obj.FileName = AFileName then
      exit;
  end;
  Result := -1;
end;


{@@ ----------------------------------------------------------------------------
  Returns the embedded object stored in the embedded object list at the
  specified index.
-------------------------------------------------------------------------------}
function TsWorkbook.GetEmbeddedObj(AIndex: Integer): TsEmbeddedObj;
begin
  Result := TsEmbeddedObj(FEmbeddedObjList[AIndex]);
end;


{@@ ----------------------------------------------------------------------------
  Returns the count of embedded objects
-------------------------------------------------------------------------------}
function TsWorkbook.GetEmbeddedObjCount: Integer;
begin
  Result := FEmbeddedObjList.Count;
end;


{@@ ----------------------------------------------------------------------------
  Returns @true if there is at least one worksheet with an embedded images.
-------------------------------------------------------------------------------}
function TsWorkbook.HasEmbeddedSheetImages: Boolean;
var
  i: Integer;
  sheet: TsWorksheet;
begin
  Result := true;
  for i:=0 to FWorksheets.Count-1 do
  begin
    sheet := TsWorksheet(FWorksheets.Items[i]);
    if sheet.GetImageCount > 0 then
      exit;
  end;
  Result := false;
end;


{@@ ----------------------------------------------------------------------------
  Removes all embedded objects
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveAllEmbeddedObj;
var
  i: Integer;
begin
  for i:= 0 to FEmbeddedObjList.Count-1 do
    TsEmbeddedObj(FEmbeddedObjList[i]).Free;
  FEmbeddedObjList.Clear;
end;


