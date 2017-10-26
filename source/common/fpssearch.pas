unit fpsSearch;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, RegExpr, fpstypes, fpspreadsheet;

type
  TsConfirmReplacementResult = (crReplace, crIgnore, crAbort);

  TsConfirmReplacementEvent = procedure (Sender: TObject; AWorksheet: TsWorksheet;
    ARow, ACol: Cardinal; const ASearchText, AReplaceText: String;
    var AResult: TsConfirmReplacementResult) of object;

  TsSearchEngine = class
  private
    FWorkbook: TsWorkbook;
    FSearchText: String;
    FSearchParams: TsSearchParams;
    FReplaceParams: TsReplaceParams;
    FRegEx: TRegExpr;
    FStopping: Boolean;
    FOnConfirmReplacement: TsConfirmReplacementEvent;
  protected
    function ExecReplace(AWorksheet: TsWorksheet; ARow, ACol: Cardinal): boolean;
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
    function FindFirst(const ASearchParams: TsSearchParams;
      out AWorksheet: TsWorksheet; out ARow, ACol: Cardinal): Boolean;
    function FindNext(const ASearchParams: TsSearchParams;
      var AWorksheet: TsWorksheet; var ARow, ACol: Cardinal): Boolean;
    function ReplaceFirst(const ASearchParams: TsSearchParams;
      const AReplaceParams: TsReplaceParams;
      out AWorksheet: TsWorksheet; out ARow, ACol: Cardinal): Boolean;
    function ReplaceNext(const ASearchParams: TsSearchParams;
      const AReplaceParams: TsReplaceParams;
      var AWorksheet: TsWorksheet; var ARow, ACol: Cardinal): Boolean;

    property OnConfirmReplacement: TsConfirmReplacementEvent
      read FOnConfirmReplacement write FOnConfirmReplacement;
  end;

implementation

uses
  lazutf8, {%H-}fpsPatches;

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

function TsSearchEngine.ExecReplace(AWorksheet: TsWorksheet; ARow, ACol: Cardinal) : Boolean;
var
  s: String;
  flags: TReplaceFlags;
  confirmation: TsConfirmReplacementResult;
begin
  if roConfirm in FReplaceParams.Options then
  begin
    if Assigned(FOnConfirmReplacement) then
    begin
      confirmation := crReplace;
      FOnConfirmReplacement(self, AWorksheet, ARow, ACol,
        FSearchParams.SearchText, FReplaceParams.ReplaceText, confirmation);
      case confirmation of
        crReplace: ;
        crIgnore : exit(false);
        crAbort  : begin FStopping := true; exit(false); end;
      end;
    end else
      raise EFPSpreadsheet.Create('[TsSearchEngine.ExecReplace] OnConfirmReplacement handler needed.');
  end;

  if roReplaceEntireCell in FReplaceParams.Options then
    AWorksheet.WriteCellValueAsString(ARow, ACol, FReplaceParams.ReplaceText)
  else begin
    s := AWorksheet.ReadAsText(ARow, ACol);
    if soCompareEntireCell in FSearchParams.Options then
      AWorksheet.WriteCellValueAsString(ARow, ACol, FReplaceParams.ReplaceText)
    else
    begin
      flags := [];
      if not (soMatchCase in FSearchParams.Options) then
        Include(flags, rfIgnoreCase);
      s := UTF8StringReplace(s, FSearchparams.SearchText, FReplaceParams.ReplaceText, flags);
      AWorksheet.WriteCellValueAsString(ARow, ACol, s);
      // to do: RegEx to be added
    end;
  end;

  Result := true;
end;

function TsSearchEngine.ExecSearch(var AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
var
  complete: boolean;
  r, c: Cardinal;
  sheet: TsWorksheet;
begin
  sheet := AWorksheet;
  r := ARow;
  c := ACol;
  complete := false;
  while (not complete) and (not Matches(AWorksheet, ARow, ACol)) do
  begin
    if soBackward in FSearchParams.Options then
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

function TsSearchEngine.FindFirst(const ASearchParams: TsSearchParams;
  out AWorksheet: TsWorksheet; out ARow, ACol: Cardinal): Boolean;
begin
  FSearchParams := ASearchParams;
  PrepareSearchText(FSearchParams.SearchText);

  if soBackward in FSearchParams.Options then
    GotoLast(AWorksheet, ARow, ACol) else
    GotoFirst(AWorksheet, ARow, ACol);

  Result := ExecSearch(AWorksheet, ARow, ACol);
end;

function TsSearchEngine.FindNext(const ASearchParams: TsSearchParams;
  var AWorksheet: TsWorksheet; var ARow, ACol: Cardinal): Boolean;
begin
  FSearchParams := ASearchParams;
  PrepareSearchText(FSearchParams.SearchText);

  if soBackward in FSearchParams.Options then
    GotoPrev(AWorksheet, ARow, ACol) else
    GotoNext(AWorksheet, ARow, ACol);

  Result := ExecSearch(AWorksheet, ARow, ACol);
end;

procedure TsSearchEngine.GotoFirst(out AWorksheet: TsWorksheet;
  out ARow, ACol: Cardinal);
begin
  if soEntireDocument in FSearchParams.Options then
    // Search entire document forward from start
    case FSearchParams.Within of
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
    if AWorksheet = nil then AWorksheet := FWorkbook.GetFirstWorksheet;
    ARow := AWorksheet.ActiveCellRow;
    ACol := AWorksheet.ActiveCellCol;
  end;
end;

procedure TsSearchEngine.GotoLast(out AWorksheet: TsWorksheet;
  out ARow, ACol: Cardinal);
begin
  if soEntireDocument in FSearchParams.Options then
    // Search entire document backward from end
    case FSearchParams.Within of
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
begin
  Result := true;

  if GotoNextInWorksheet(AWorksheet, ARow, ACol) then
    exit;

  case FSearchParams.Within of
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
        if (soWrapDocument in FSearchParams.Options) then
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(0);
          ARow := 0;
          ACol := 0;
          exit;
        end;
      end;

    swWorksheet:
      if soWrapDocument in FSearchParams.Options then begin
        ARow := 0;
        ACol := 0;
        exit;
      end;

    swColumn:
      if soWrapDocument in FSearchParams.Options then begin
        ARow := 0;
        ACol := AWorksheet.ActiveCellCol;
        exit;
      end;

    swRow:
      if soWrapDocument in FSearchParams.Options then begin
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
  if (soAlongRows in FSearchParams.Options) or (FSearchParams.Within = swRow) then
  begin
    inc(ACol);
    if ACol <= AWorksheet.GetLastColIndex then
      exit;
    if (FSearchParams.Within <> swRow) then
    begin
      ACol := 0;
      inc(ARow);
      if ARow <= AWorksheet.GetLastRowIndex then
        exit;
    end;
  end else
  if not (soAlongRows in FSearchParams.Options) or (FSearchParams.Within = swColumn) then
  begin
    inc(ARow);
    if ARow <= AWorksheet.GetLastRowIndex then
      exit;
    if (FSearchParams.Within <> swColumn) then
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
begin
  Result := true;

  if GotoPrevInWorksheet(AWorksheet, ARow, ACol) then
    exit;

  case FSearchParams.Within of
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
        if (soWrapDocument in FSearchParams.Options) then
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(FWorkbook.GetWorksheetCount-1);
          ARow := AWorksheet.GetLastRowIndex;
          ACol := AWorksheet.GetLastColIndex;
          exit;
        end;
      end;

    swWorksheet:
      if soWrapDocument in FSearchParams.Options then
      begin
        ARow := AWorksheet.GetLastRowIndex;
        ACol := AWorksheet.GetLastColIndex;
        exit;
      end;

    swColumn:
      if soWrapDocument in FSearchParams.Options then
      begin
        ARow := AWorksheet.GetLastRowIndex;
        ACol := AWorksheet.ActiveCellCol;
        exit;
      end;

    swRow:
      if soWrapDocument in FSearchParams.Options then
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
  if (soAlongRows in FSearchParams.Options) or (FSearchParams.Within = swRow) then
  begin
    if ACol > 0 then begin
      dec(ACol);
      exit;
    end;
    if (FSearchParams.Within <> swRow) then
    begin
      ACol := AWorksheet.GetLastColIndex;
      if ARow > 0 then
      begin
        dec(ARow);
        exit;
      end;
    end;
  end else
  if not (soAlongRows in FSearchParams.Options) or (FSearchParams.Within = swColumn) then
  begin
    if ARow > 0 then begin
      dec(ARow);
      exit;
    end;
    if (FSearchParams.Within <> swColumn) then
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
  txt: String;
begin
  cell := AWorksheet.FindCell(ARow, ACol);
  txt := '';
  if cell <> nil then
  begin
    if (soSearchInComment in FSearchParams.Options) then
      txt := AWorksheet.ReadComment(cell)
    else
      txt := AWorksheet.ReadAsText(cell);
  end;

  if soRegularExpr in FSearchParams.Options then
    Result := FRegEx.Exec(txt)
  else
  begin
    if not (soMatchCase in FSearchParams.Options) then
      txt := UTF8Lowercase(txt);
    if soCompareEntireCell in FSearchParams.Options then
      exit(txt = FSearchText);
    if pos(FSearchText, txt) > 0 then
      exit(true);
    Result := false;
  end;
end;

procedure TsSearchEngine.PrepareSearchText(const ASearchText: String);
begin
  if soRegularExpr in FSearchParams.Options then
  begin
    FreeAndNil(FRegEx);
    FRegEx := TRegExpr.Create;
    FRegEx.Expression := ASearchText
  end else
  if (soMatchCase in FSearchParams.Options) then
    FSearchText := ASearchText else
    FSearchText := UTF8Lowercase(ASearchText);
end;

function TsSearchEngine.ReplaceFirst(const ASearchParams: TsSearchParams;
  const AReplaceParams: TsReplaceParams; out AWorksheet: TsWorksheet;
  out ARow, ACol: Cardinal): Boolean;
var
  r,c: Cardinal;
  sheet: TsWorksheet;
begin
  FStopping := false;

  // Lock the visual components in case of "replace all" and "no confirmation"
  if AReplaceParams.Options * [roReplaceAll, roConfirm] = [roReplaceAll] then
    FWorkbook.DisableNotifications;

  Result := FindFirst(ASearchParams, AWorksheet, ARow, ACol);
  r := ARow;
  c := ACol;
  sheet := AWorksheet;

  if Result then
  begin
    FReplaceParams := AReplaceParams;
    Result := ExecReplace(AWorksheet, ARow, ACol);
    if roReplaceAll in FReplaceParams.Options then
    begin
      while (not FStopping) and FindNext(FSearchParams, AWorksheet, ARow, ACol) do
      begin
        r := ARow;
        c := ACol;
        sheet := AWorksheet;
        ExecReplace(AWorksheet, ARow, ACol);
      end;
    end;
  end;

  // Unlock the visual components in case of "replace all" and "no confirmation"
  // and select the last replaced cell
  if AReplaceParams.Options * [roReplaceAll, roConfirm] = [roReplaceAll] then
  begin
    FWorkbook.EnableNotifications;
    if Result then
    begin
      FWorkbook.SelectWorksheet(sheet);
      sheet.SelectCell(r, c);
    end;
  end;
end;

function TsSearchEngine.ReplaceNext(const ASearchParams: TsSearchParams;
  const AReplaceParams: TsReplaceParams; var AWorksheet: TsWorksheet;
  var ARow, ACol: Cardinal): Boolean;
var
  r, c: Cardinal;
  sheet: TsWorksheet;
begin
  FStopping := false;

  // Lock the visual components in case of "replace all" and "no confirmation"
  if AReplaceParams.Options * [roReplaceAll, roConfirm] = [roReplaceAll] then
    FWorkbook.DisableNotifications;

  Result := FindNext(ASearchParams, AWorksheet, ARow, ACol);
  r := ARow;
  c := ACol;
  sheet := AWorksheet;

  if Result then
  begin
    FReplaceParams := AReplaceParams;
    Result := ExecReplace(AWorksheet, ARow, ACol);
    if roReplaceAll in FReplaceParams.Options then
    begin
      while (not FStopping) and FindNext(FSearchParams, AWorksheet, ARow, ACol) do
      begin
        r := ARow;
        c := ACol;
        sheet := AWorksheet;
        ExecReplace(AWorksheet, ARow, ACol);
      end;
    end;
  end;

  // Unlock the visual components in case of "replace all" and "no confirmation"
  // and select the last replaced cell
  if AReplaceParams.Options * [roReplaceAll, roConfirm] = [roReplaceAll] then
  begin
    FWorkbook.EnableNotifications;
    if Result then
    begin
      FWorkbook.SelectWorksheet(sheet);
      sheet.SelectCell(r, c);
    end;
  end;
end;

end.

