unit sSearchForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ExtCtrls, Buttons, fpsTypes, fpspreadsheet, fpsSearch;

type
  TsSearchEvent = procedure (Sender: TObject; AFound: Boolean;
    AWorksheet: TsWorksheet; ARow, ACol: Cardinal) of object;

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
    RgSearchWithin: TRadioGroup;
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SearchButtonClick(Sender: TObject);
  private
    { private declarations }
    FSearchEngine: TsSearchEngine;
    FWorkbook: TsWorkbook;
    FFoundWorksheet: TsWorksheet;
    FFoundRow, FFoundCol: Cardinal;
    FOnFound: TsSearchEvent;
    function GetParams: TsSearchParams;
    procedure SetParams(const AValue: TsSearchParams);
  public
    { public declarations }
    procedure Execute(AWorkbook: TsWorkbook);
    property Workbook: TsWorkbook read FWorkbook;
    property SearchParams: TsSearchParams read GetParams write SetParams;
    property OnFound: TsSearchEvent read FOnFound write FOnFound;
  end;

var
  SearchForm: TSearchForm;

  DefaultSearchParams: TsSearchParams = (
    SearchText: '';
    Options: [];
    Within: swWorksheet
  );


implementation

{$R *.lfm}

uses
  fpsUtils;

const
  MAX_SEARCH_ITEMS      = 10;

  COMPARE_ENTIRE_CELL   = 0;
  MATCH_CASE            = 1;
  REGULAR_EXPRESSION    = 2;
  SEARCH_ALONG_ROWS     = 3;
  CONTINUE_AT_START_END = 4;


{ TSearchForms }

procedure TSearchForm.Execute(AWorkbook: TsWorkbook);
begin
  FWorkbook := AWorkbook;
  Show;
end;

procedure TSearchForm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  P: TPoint;
begin
  Unused(CloseAction);
  FreeAndNil(FSearchEngine);
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

  FFoundCol := UNASSIGNED_ROW_COL_INDEX;
  FFoundRow := UNASSIGNED_ROW_COL_INDEX;
  FFoundWorksheet := nil;
end;

function TSearchForm.GetParams: TsSearchParams;
begin
  Result.SearchText := CbSearchText.Text;
  Result.Options := [];
  if CgSearchOptions.Checked[COMPARE_ENTIRE_CELL] then
    Include(Result.Options, soCompareEntireCell);
  if CgSearchOptions.Checked[MATCH_CASE] then
    Include(Result.Options, soMatchCase);
  if CgSearchOptions.Checked[REGULAR_EXPRESSION] then
    Include(Result.Options, soRegularExpr);
  if CgSearchOptions.Checked[SEARCH_ALONG_ROWS] then
    Include(Result.Options, soAlongRows);
  if CgSearchOptions.Checked[CONTINUE_AT_START_END] then
    Include(Result.Options, soWrapDocument);
  if RgSearchStart.ItemIndex = 1 then
    Include(Result.Options, soEntireDocument);
  Result.Within := TsSearchWithin(RgSearchWithin.ItemIndex);
end;

procedure TSearchForm.SearchButtonClick(Sender: TObject);
var
  params: TsSearchParams;
  found: Boolean;
begin
  params := GetParams;
  if params.SearchText = '' then
    exit;

  if CbSearchText.Items.IndexOf(params.SearchText) = -1 then
  begin
    CbSearchText.Items.Insert(0, params.SearchText);
    while CbSearchText.Items.Count > MAX_SEARCH_ITEMS do
      CbSearchText.Items.Delete(CbSearchText.Items.Count-1);
  end;

  if FSearchEngine = nil then
  begin
    FSearchEngine := TsSearchEngine.Create(FWorkbook);
    if (soBackward in params.Options) then
      Include(params.Options, soBackward) else
      Exclude(params.Options, soBackward);
    found := FSearchEngine.FindFirst(params.SearchText, params, FFoundWorksheet, FFoundRow, FFoundCol);
  end else
  begin
    if (Sender = BtnSearchBack) then
      Include(params.Options, soBackward) else
      Exclude(params.Options, soBackward);
    // User may select a different worksheet/different cell to continue search!
    FFoundWorksheet := FWorkbook.ActiveWorksheet;
    FFoundRow := FFoundWorksheet.ActiveCellRow;
    FFoundCol := FFoundWorksheet.ActiveCellCol;
    found := FSearchEngine.FindNext(params.SearchText, params, FFoundWorksheet, FFoundRow, FFoundCol);
  end;

  if Assigned(FOnFound) then
    FOnFound(self, found, FFoundWorksheet, FFoundRow, FFoundCol);

  BtnSearchBack.Visible := true;
  BtnSearch.Caption := 'Next';
end;

procedure TSearchForm.SetParams(const AValue: TsSearchParams);
begin
  CbSearchText.Text := Avalue.SearchText;
  CgSearchOptions.Checked[COMPARE_ENTIRE_CELL] := (soCompareEntireCell in AValue.Options);
  CgSearchOptions.Checked[MATCH_CASE] := (soMatchCase in AValue.Options);
  CgSearchOptions.Checked[REGULAR_EXPRESSION] := (soRegularExpr in Avalue.Options);
  CgSearchOptions.Checked[SEARCH_ALONG_ROWS] := (soAlongRows in AValue.Options);
  CgSearchOptions.Checked[CONTINUE_AT_START_END] := (soWrapDocument in Avalue.Options);
  RgSearchWithin.ItemIndex := ord(AValue.Within);
  RgSearchStart.ItemIndex := ord(soEntireDocument in AValue.Options);
end;

end.

