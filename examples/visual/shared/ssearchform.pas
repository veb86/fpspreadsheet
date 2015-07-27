unit sSearchForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ButtonPanel,
  StdCtrls, ExtCtrls, Buttons, fpsTypes, fpspreadsheet;

type

  { TSearchParams }

  TsSearchSource = (spsWorksheet, spsWorkbook);
  TsSearchStart  = (spsBeginningEnd, spsActiveCell);

  TsSearchParams = record
    SearchText: String;
    Options: TsSearchOptions;
    Source: TsSearchSource;
    Start: TsSearchStart;
  end;

  TsSearchEvent = procedure (Sender: TObject; ACell: PCell) of object;


  { TSearchForm }

  TSearchForm = class(TForm)
    Bevel1: TBevel;
    BtnSearchBack: TBitBtn;
    BtnClose: TBitBtn;
    BtnSearch: TBitBtn;
    CbSearchText: TComboBox;
    CgSearchOptions: TCheckGroup;
    LblSearchText: TLabel;
    ButtonPanel: TPanel;
    RgSearchStart: TRadioGroup;
    RgSearchSource: TRadioGroup;
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SearchButtonClick(Sender: TObject);
  private
    { private declarations }
    FWorkbook: TsWorkbook;
    FFoundCell: PCell;
    FOnFound: TsSearchEvent;
    procedure CtrlsToParams(var ASearchParams: TsSearchParams);
    function FindStartCell(AParams: TsSearchParams; var AWorksheet: TsWorksheet;
      var AStartRow, AStartCol: Cardinal): Boolean;
    procedure ParamsToCtrls(const ASearchParams: TsSearchParams);
  public
    { public declarations }
    procedure Execute(AWorkbook: TsWorkbook; var ASearchParams: TsSearchParams);
    property Workbook: TsWorkbook read FWorkbook;
    property OnFound: TsSearchEvent read FOnFound write FOnFound;
  end;

var
  SearchForm: TSearchForm;

  DefaultSearchParams: TsSearchParams = (
    SearchText: '';
    Options: [soIgnoreCase];
    Source: spsWorksheet;
    Start: spsActiveCell;
  );


implementation

{$R *.lfm}

const
  MAX_SEARCH_ITEMS = 10;

procedure TSearchForm.CtrlsToParams(var ASearchParams: TsSearchParams);
var
  i: Integer;
begin
  ASearchParams.SearchText := CbSearchText.Text;
  ASearchParams.Options := [];
  for i:=0 to CgSearchOptions.Items.Count-1 do
    if CgSearchOptions.Checked[i] then
      Include(ASearchparams.Options, TsSearchOption(i));
  ASearchParams.Source := TsSearchSource(RgSearchSource.ItemIndex);
  ASearchParams.Start := TsSearchStart(RgSearchStart.ItemIndex);
end;

procedure TSearchForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  P: TPoint;
begin
  P.X := Left;
  P.Y := Top;
  Position := poDesigned;
  Left := P.X;
  Top := P.Y;
end;

procedure TSearchForm.FormCreate(Sender: TObject);
begin
  Position := poMainFormCenter;
end;

procedure TSearchForm.FormShow(Sender: TObject);
begin
  BtnSearch.Caption := 'Search';
  BtnSearchBack.Visible := false;
  FFoundCell := nil;
end;

procedure TSearchForm.Execute(AWorkbook: TsWorkbook;
  var ASearchParams: TsSearchParams);
begin
  FWorkbook := AWorkbook;
  ParamsToCtrls(ASearchParams);
  Show;
  CtrlsToParams(ASearchParams);
end;

function TSearchForm.FindStartCell(AParams: TsSearchParams;
  var AWorksheet: TsWorksheet; var AStartRow, AStartCol: Cardinal): Boolean;
var
  sheetIndex: integer;
  cell: PCell;
begin
  Result := false;
  cell := nil;

  // Case (1): Search not executed before
  if FFoundCell = nil then
  begin
    case AParams.Start of
      spsActiveCell:
        begin
          AWorksheet := FWorkbook.ActiveWorksheet;
          AStartRow := AWorksheet.ActiveCellRow;
          AStartCol := AWorksheet.ActiveCellCol;
        end;
      spsBeginningEnd:
        if (soBackward in AParams.Options) then
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(FWorkbook.GetWorksheetCount-1);
          AStartCol := AWorksheet.GetLastColIndex;
          AStartRow := AWorksheet.GetlastRowIndex;
        end else
        begin
          AWorksheet := FWorkbook.GetWorksheetByIndex(0);
          AStartCol := AWorksheet.GetFirstColIndex;
          AStartRow := AWorksheet.GetFirstRowIndex;
        end;
    end;
  end else
  // Case (2):
  // Repeated execution of search to start at cell adjacent to the one found in
  // previous call.
  begin
    //AWorksheet := TsWorksheet(FFoundCell^.Worksheet);
    // FoundCell is the cell found in the previous call.
    //AStartRow := FFoundCell^.Row;
    //AStartCol := FFoundCell^.Col;
    sheetIndex := FWorkbook.GetWorksheetIndex(AWorksheet);
    // Case (1): Find prior occupied cell along row
    if (AParams.Options * [soAlongRows, soBackward] = [soAlongRows, soBackward]) then
    begin
      cell := AWorksheet.FindPrevCellInRow(AStartRow, AStartCol);
      // No "prior" cell found in this row --> Proceed with previous row
      while (cell = nil) and (AStartRow > 0) do
      begin
        dec(AStartRow);
        AStartCol := AWorksheet.GetLastColIndex;
        cell := AWorksheet.FindCell(AStartRow, AStartCol);
        if (cell = nil) then
          cell := AWorksheet.FindPrevCellInRow(AStartRow, AStartCol);
        // No "prior" cell found in this sheet --> Proceed with previous sheet
        if (cell = nil) and (AStartRow = 0) then
        begin
          if sheetIndex = 0 then
            exit;
          dec(sheetIndex);
          AWorksheet := FWorkbook.GetWorksheetByIndex(sheetIndex);
          AStartCol := AWorksheet.GetLastColIndex;
          AStartRow := AWorksheet.GetLastRowIndex;
          cell := AWorksheet.FindCell(AStartRow, AStartCol);
          if (cell = nil) then
            cell := AWorksheet.FindPrevCellInRow(AStartRow, AStartCol);
        end;
      end;
    end
    else
    // Case (2): Find prior occupied cell along columns
    if (AParams.Options * [soAlongRows, soBackward] = [soBackward]) then
    begin
      cell := AWorksheet.FindPrevCellInCol(AStartRow, AStartCol);
      // No "preior" cell found in this column --> Proceed with previous column
      while (cell = nil) and (AStartCol > 0) do
      begin
        dec(AStartCol);
        AStartRow := AWorksheet.GetLastRowIndex;
        cell := AWorksheet.FindCell(AStartRow, AStartCol);
        if (cell = nil) then
          cell := AWorksheet.FindPrevCellInCol(AStartRow, AStartCol);
        // No "prior" cell found in this sheet --> Proceed with previous sheet
        if (cell = nil) and (AStartCol = 0) then
        begin
          if sheetIndex = 0 then
            exit;
          dec(sheetIndex);
          AWorksheet := FWorkbook.GetWorksheetByIndex(sheetIndex);
          AStartCol := AWorksheet.GetLastColIndex;
          AStartRow := AWorksheet.GetLastRowIndex;
          cell := AWorksheet.FindCell(AStartRow, AStartCol);
          if (cell = nil) then
            cell := AWorksheet.FindPrevCellinCol(AStartRow, AStartCol);
        end;
      end;
    end
    else
    // Case (3): Find next occupied cell along row
    if (AParams.Options * [soAlongRows, soBackward] = [soAlongRows]) then
    begin
      cell := AWorksheet.FindNextCellInRow(AStartRow, AStartCol);
      // No cell found in this row --> Proceed with next row
      while (cell = nil) and (AStartRow < AWorksheet.GetLastRowIndex) do
      begin
        inc(AStartRow);
        AStartCol := AWorksheet.GetFirstColIndex;
        cell := AWorksheet.FindCell(AStartRow, AStartCol);
        if (cell = nil) then
          cell := AWorksheet.FindNextCellInRow(AStartRow, AStartCol);
        // No "next" cell found in this sheet --> Proceed with next sheet
        if (cell = nil) and (AStartRow = AWorksheet.GetLastRowIndex) then
        begin
          if sheetIndex = 0 then
            exit;
          inc(sheetIndex);
          AWorksheet := FWorkbook.GetWorksheetByIndex(sheetIndex);
          AStartCol := AWorksheet.GetLastColIndex;
          AStartRow := AWorksheet.GetLastRowIndex;
          cell := AWorksheet.FindCell(AStartRow, AStartCol);
          if (cell = nil) then
            cell := AWorksheet.FindNextCellInRow(AStartRow, AStartCol);
        end;
      end;
    end
    else
    // Case (4): Find next occupied cell along column
    if (AParams.Options * [soAlongRows, soBackward] = []) then
    begin
      cell := AWorksheet.FindNextCellInCol(AStartRow, AStartCol);
      // No "next" occupied cell found in this column --> Proceed with next column
      while (cell = nil) and (AStartCol < AWorksheet.GetLastColIndex) do
      begin
        inc(AStartCol);
        AStartRow := AWorksheet.GetFirstRowIndex;
        cell := AWorksheet.FindCell(AStartRow, AStartCol);
        if (cell = nil) then
          cell := AWorksheet.FindNextCellInCol(AStartRow, AStartCol);
        // No "next" cell found in this sheet --> Proceed with next sheet
        if (cell = nil) and (AStartCol = 0) then
        begin
          if sheetIndex = 0 then
            exit;
          inc(sheetIndex);
          AWorksheet := FWorkbook.GetWorksheetByIndex(sheetIndex);
          AStartCol := AWorksheet.GetLastColIndex;
          AStartRow := AWorksheet.GetLastRowIndex;
          cell := AWorksheet.FindCell(AStartRow, AStartCol);
          if (cell = nil) then
            cell := AWorksheet.FindNextCellInCol(AStartRow, AStartCol);
        end;
      end;
    end;
  end;
  if cell <> nil then
  begin
    AStartRow := cell^.Row;
    AStartCol := cell^.Col;
  end;
  Result := true;
end;

procedure TSearchForm.ParamsToCtrls(const ASearchParams: TsSearchParams);
var
  i: Integer;
  o: TsSearchOption;
begin
  CbSearchText.Text := ASearchParams.SearchText;
  for o in TsSearchOption do
    if ord(o) < CgSearchOptions.Items.Count then
      CgSearchOptions.Checked[ord(o)] := (o in ASearchParams.Options);
  RgSearchSource.ItemIndex := ord(ASearchParams.Source);
  RgSearchStart.ItemIndex := ord(ASearchParams.Start);
end;

procedure TSearchForm.SearchButtonClick(Sender: TObject);
var
  startsheet: TsWorksheet;
  sheetIdx: Integer;
  r,c: Cardinal;
  backward: Boolean;
  params: TsSearchParams;
  cell: PCell;
begin
  CtrlsToParams(params);
  if params.SearchText = '' then
    exit;

  if CbSearchText.Items.IndexOf(params.SearchText) = -1 then
  begin
    CbSearchText.Items.Insert(0, params.SearchText);
    while CbSearchText.Items.Count > MAX_SEARCH_ITEMS do
      CbSearchText.Items.Delete(CbSearchText.Items.Count-1);
  end;

  if FFoundcell = nil then
    backward := (soBackward in params.Options)  // 1st call: use value from Options
  else
    backward := (Sender = BtnSearchBack);       // subseq call: follow button
  if backward then Include(params.Options, soBackward) else
    Exclude(params.Options, soBackward);

  if params.Start = spsActiveCell then
  begin
    startSheet := FWorkbook.ActiveWorksheet;
    FFoundCell := startSheet.FindCell(startSheet.ActiveCellRow, startSheet.ActiveCellCol);
  end;

  if FFoundCell <> nil then
  begin
    startsheet := TsWorksheet(FFoundCell^.Worksheet);
    r := FFoundCell^.Row;
    c := FFoundCell^.Col;
  end;
  cell := nil;

  while FindStartCell(params, startsheet, r, c) and (cell = nil) do
  begin
    cell := startsheet.Search(params.SearchText, params.Options, r, c);
    if (cell <> nil) then
    begin
      FWorkbook.SelectWorksheet(startsheet);
      startsheet.SelectCell(cell^.Row, cell^.Col);
      if Assigned(FOnFound) then FOnFound(Self, cell);
      FFoundCell := cell;
      break;
    end;
    // not found --> go to next sheet
    case params.Source of
      spsWorksheet:
        break;
      spsWorkbook:
        begin
          sheetIdx := FWorkbook.GetWorksheetIndex(startsheet);
          if backward then
          begin
            if (sheetIdx = 0) then exit;
            startsheet := FWorkbook.GetWorksheetByIndex(sheetIdx-1);
            r := startsheet.GetLastRowIndex;
            c := startsheet.GetLastColIndex;
          end else
          begin
            if (sheetIdx = FWorkbook.GetWorksheetCount-1) then exit;
            startsheet := FWorkbook.GetWorksheetByIndex(sheetIdx+1);
            r := 0;
            c := 0;
          end;
        end;
    end;
  end;

  BtnSearchBack.Visible := true;
  BtnSearch.Caption := 'Next';
end;


end.

