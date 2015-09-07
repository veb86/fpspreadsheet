unit fpsSearch;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, RegExpr, fpstypes, fpspreadsheet;

type
  TsSearchEngine = class
  private
    FWorkbook: TsWorkbook;
    FSearchText: String;
    FParams: TsSearchParams;
    FCurrSel: Integer;
    FRegEx: TRegExpr;
  protected
    function ExecSearch(var AWorksheet: TsWorksheet;
      var ARow, ACol: Cardinal): Boolean;
    procedure GotoFirst(out AWorksheet: TsWorksheet; out ARow, ACol: Cardinal);
    procedure GotoLast(out AWorksheet: TsWorksheet; out ARow, ACol: Cardinal);
    function GotoNext(var AWorksheet: TsWorksheet;
      var ARow, ACol: Cardinal): Boolean;
    function GotoNextInWorksheet(AWorksheet: TsWorksheet;
      var ARow, ACol: Cardinal): Boolean;
    function GotoPrev(var AWorksheet: TsWorksheet;
      var ARow, ACol: Cardinal): Boolean;
    function GotoPrevInWorksheet(AWorksheet: TsWorksheet;
      var ARow, ACol: Cardinal): Boolean;
    function Matches(AWorksheet: TsWorksheet; ARow, ACol: Cardinal): Boolean;
    procedure PrepareSearchText(const ASearchText: String);

  public
    constructor Create(AWorkbook: TsWorkbook);
    destructor Destroy; override;
    function FindFirst(const ASearchText: String; const AParams: TsSearchParams;
      out AWorksheet: TsWorksheet; out ARow, ACol: Cardinal): Boolean;
    function FindNext(const ASearchText: String; const AParams: TsSearchParams;
      var AWorksheet: TsWorksheet; var ARow, ACol: Cardinal): Boolean;
  end;

implementation

uses
  lazutf8;

constructor TsSearchEngine.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
end;

destructor TsSearchEngine.Destroy;
begin
  FreeAndNil(FRegEx);
  inherited Destroy;
end;

function TsSearchEngine.ExecSearch(var AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
var
  complete: boolean;
  r, c: LongInt;
  sheet: TsWorksheet;
begin
  sheet := AWorksheet;
  r := ARow;
  c := ACol;
  complete := false;
  while (not complete) and (not Matches(AWorksheet, ARow, ACol)) do
  begin
    if soBackward in FParams.Options then
      complete := not GotoPrev(AWorkSheet, ARow, ACol) else
      complete := not GotoNext(AWorkSheet, ARow, ACol);
    // Avoid infinite loop if search phrase does not exist in document.
    if (AWorksheet = sheet) and (ARow = r) and (ACol = c) then
      complete := true;
  end;
  Result := not complete;
  if Result then
  begin
    FWorkbook.SelectWorksheet(AWorksheet);
    AWorksheet.SelectCell(ARow, ACol);
  end else
  begin
    AWorksheet := nil;
    ARow := UNASSIGNED_ROW_COL_INDEX;
    ACol := UNASSIGNED_ROW_COL_INDEX;
  end;
end;

function TsSearchEngine.FindFirst(const ASearchText: String;
  const AParams: TsSearchParams; out AWorksheet: TsWorksheet;
  out ARow, ACol: Cardinal): Boolean;
begin
  FParams := AParams;
  PrepareSearchText(ASearchText);

  if soBackward in FParams.Options then
    GotoLast(AWorksheet, ARow, ACol) else
    GotoFirst(AWorksheet, ARow, ACol);

  Result := ExecSearch(AWorksheet, ARow, ACol);
end;

function TsSearchEngine.FindNext(const ASearchText: String;
  const AParams: TsSearchParams; var AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
begin
  FParams := AParams;
  PrepareSearchText(ASearchText);

  if soBackward in FParams.Options then
    GotoPrev(AWorksheet, ARow, ACol) else
    GotoNext(AWorksheet, ARow, ACol);

  Result := ExecSearch(AWorksheet, ARow, ACol);
end;

procedure TsSearchEngine.GotoFirst(out AWorksheet: TsWorksheet;
  out ARow, ACol: Cardinal);
begin
  if soEntireDocument in FParams.Options then
    // Search entire document forward from start
    case FParams.Within of
      swWorkbook :
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(0);
          ARow := 0;
          ACol := 0;
        end;
      swWorksheet:
        begin
          AWorksheet := FWorkbook.ActiveWorksheet;
          ARow := 0;
          ACol := 0;
        end;
      swColumn:
        begin
          AWorksheet := FWorkbook.ActiveWorksheet;
          ARow := 0;
          ACol := AWorksheet.ActiveCellCol;
        end;
      swRow:
        begin
          AWorksheet := FWorkbook.ActiveWorksheet;
          ARow := AWorksheet.ActiveCellRow;
          ACol := 0;
        end;
    end
  else
  begin
    // Search starts at active cell
    AWorksheet := FWorkbook.ActiveWorksheet;
    ARow := AWorksheet.ActiveCellRow;
    ACol := AWorksheet.ActiveCellCol;
  end;
end;

procedure TsSearchEngine.GotoLast(out AWorksheet: TsWorksheet;
  out ARow, ACol: Cardinal);
var
  cell: PCell;
  sel: TsCellRangeArray;
begin
  if soEntireDocument in FParams.Options then
    // Search entire document backward from end
    case FParams.Within of
      swWorkbook :
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(FWorkbook.GetWorksheetCount-1);
          ARow := AWorksheet.GetLastRowIndex;
          ACol := AWorksheet.GetLastColIndex;
        end;
      swWorksheet:
        begin
          AWorksheet := FWorkbook.ActiveWorksheet;
          ARow := AWorksheet.GetLastRowIndex;
          ACol := AWorksheet.GetLastColIndex;
        end;
      swColumn:
        begin
          AWorksheet := FWorkbook.ActiveWorksheet;
          ARow := AWorksheet.GetLastRowIndex;
          ACol := AWorksheet.ActiveCellCol;
        end;
      swRow:
        begin
          AWorksheet := FWorkbook.ActiveWorksheet;
          ARow := AWorksheet.ActiveCellRow;
          ACol := AWorksheet.GetLastColIndex;
        end;
    end
  else
  begin
    // Search starts at active cell
    AWorksheet := FWorkbook.ActiveWorksheet;
    ARow := AWorksheet.ActiveCellRow;
    ACol := AWorksheet.ActiveCellCol;
  end;
end;

function TsSearchEngine.GotoNext(var AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
var
  idx: Integer;
  sel: TsCellRangeArray;
begin
  Result := true;

  if GotoNextInWorksheet(AWorksheet, ARow, ACol) then
    exit;

  case FParams.Within of
    swWorkbook:
      begin
        // Need to go to next sheet
        idx := FWorkbook.GetWorksheetIndex(AWorksheet) + 1;
        if idx < FWorkbook.GetWorksheetCount then
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(idx);
          ARow := 0;
          ACol := 0;
          exit;
        end;
        // Continue search with first worksheet
        if (soWrapDocument in FParams.Options) then
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(0);
          ARow := 0;
          ACol := 0;
          exit;
        end;
      end;

    swWorksheet:
      if soWrapDocument in FParams.Options then begin
        ARow := 0;
        ACol := 0;
        exit;
      end;

    swColumn:
      if soWrapDocument in FParams.Options then begin
        ARow := 0;
        ACol := AWorksheet.ActiveCellCol;
        exit;
      end;

    swRow:
      if soWrapDocument in FParams.Options then begin
        ARow := AWorksheet.ActiveCellRow;
        ACol := 0;
        exit;
      end;
  end;  // case

  Result := false;
end;


function TsSearchEngine.GotoNextInWorksheet(AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
begin
  Result := true;
  if (soAlongRows in FParams.Options) or (FParams.Within = swRow) then
  begin
    inc(ACol);
    if ACol <= AWorksheet.GetLastColIndex then
      exit;
    if (FParams.Within <> swRow) then
    begin
      ACol := 0;
      inc(ARow);
      if ARow <= AWorksheet.GetLastRowIndex then
        exit;
    end;
  end else
  if not (soAlongRows in FParams.Options) or (FParams.Within = swColumn) then
  begin
    inc(ARow);
    if ARow <= AWorksheet.GetLastRowIndex then
      exit;
    if (FParams.Within <> swColumn) then
    begin
      ARow := 0;
      inc(ACol);
      if (ACol <= AWorksheet.GetLastColIndex) then
        exit;
    end;
  end;
  // We reached the last cell, there is no "next" cell in this sheet
  Result := false;
end;

function TsSearchEngine.GotoPrev(var AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
var
  idx: Integer;
  sel: TsCellRangeArray;
begin
  Result := true;

  if GotoPrevInWorksheet(AWorksheet, ARow, ACol) then
    exit;

  case FParams.Within of
    swWorkbook:
      begin
        // Need to go to previous sheet
        idx := FWorkbook.GetWorksheetIndex(AWorksheet) - 1;
        if idx >= 0 then
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(idx);
          ARow := AWorksheet.GetLastRowIndex;
          ACol := AWorksheet.GetlastColIndex;
          exit;
        end;
        if (soWrapDocument in FParams.Options) then
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(FWorkbook.GetWorksheetCount-1);
          ARow := AWorksheet.GetLastRowIndex;
          ACol := AWorksheet.GetLastColIndex;
          exit;
        end;
      end;

    swWorksheet:
      if soWrapDocument in FParams.Options then
      begin
        ARow := AWorksheet.GetLastRowIndex;
        ACol := AWorksheet.GetLastColIndex;
        exit;
      end;

    swColumn:
      if soWrapDocument in FParams.Options then
      begin
        ARow := AWorksheet.GetLastRowIndex;
        ACol := AWorksheet.ActiveCellCol;
        exit;
      end;

    swRow:
      if soWrapDocument in FParams.Options then
      begin
        ARow := AWorksheet.ActiveCellRow;
        ACol := AWorksheet.GetLastColIndex;
        exit;
      end;
  end;  // case

  Result := false;
end;

function TsSearchEngine.GotoPrevInWorksheet(AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
begin
  Result := true;
  if (soAlongRows in FParams.Options) or (FParams.Within = swRow) then
  begin
    if ACol > 0 then begin
      dec(ACol);
      exit;
    end;
    if (FParams.Within <> swRow) then
    begin
      ACol := AWorksheet.GetLastColIndex;
      if ARow > 0 then
      begin
        dec(ARow);
        exit;
      end;
    end;
  end else
  if not (soAlongRows in FParams.Options) or (FParams.Within = swColumn) then
  begin
    if ARow > 0 then begin
      dec(ARow);
      exit;
    end;
    if (FParams.Within <> swColumn) then
    begin
      ARow := AWorksheet.GetlastRowIndex;
      if ACol > 0 then
      begin
        dec(ACol);
        exit;
      end;
    end;
  end;
  // We reached the first cell, there is no "previous" cell
  Result := false;
end;

function TsSearchEngine.Matches(AWorksheet: TsWorksheet; ARow, ACol: Cardinal): Boolean;
var
  cell: PCell;
  celltxt: String;
begin
  cell := AWorksheet.FindCell(ARow, ACol);
  if cell <> nil then
    celltxt := AWorksheet.ReadAsText(cell) else
    celltxt := '';

  if soRegularExpr in FParams.Options then
    Result := FRegEx.Exec(celltxt)
  else
  begin
    if not (soMatchCase in FParams.Options) then
      celltxt := UTF8Lowercase(celltxt);
    if soCompareEntireCell in FParams.Options then
      exit(celltxt = FSearchText);
    if UTF8Pos(FSearchText, celltxt) > 0 then
      exit(true);
    Result := false;
  end;
end;

procedure TsSearchEngine.PrepareSearchText(const ASearchText: String);
begin
  if soRegularExpr in FParams.Options then
  begin
    FreeAndNil(FRegEx);
    FRegEx := TRegExpr.Create;
    FRegEx.Expression := ASearchText
  end else
  if (soMatchCase in FParams.Options) then
    FSearchText := ASearchText else
    FSearchText := UTF8Lowercase(ASearchText);
end;

end.

