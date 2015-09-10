unit sSearchForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ExtCtrls, Buttons, ComCtrls, fpsTypes, fpspreadsheet, fpsSearch;

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
    CbReplaceText: TComboBox;
    CgOptions: TCheckGroup;
    LblSearchText: TLabel;
    ButtonPanel: TPanel;
    LblSearchText1: TLabel;
    SearchParamsPanel: TPanel;
    SearchTextPanel: TPanel;
    RgSearchStart: TRadioGroup;
    RgSearchWithin: TRadioGroup;
    ReplaceTextPanel: TPanel;
    TabControl: TTabControl;
    procedure ExecuteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
    procedure TabControlChanging(Sender: TObject; var AllowChange: Boolean);
  private
    { private declarations }
    FSearchEngine: TsSearchEngine;
    FWorkbook: TsWorkbook;
    FFoundWorksheet: TsWorksheet;
    FFoundRow, FFoundCol: Cardinal;
    FSearchParams: TsSearchParams;
    FReplaceParams: TsReplaceParams;
    FOnFound: TsSearchEvent;
    function GetReplaceParams: TsReplaceParams;
    function GetSearchParams: TsSearchParams;
    procedure SetReplaceParams(const AValue: TsReplaceParams);
    procedure SetSearchParams(const AValue: TsSearchParams);
  protected
    procedure ConfirmReplacementHandler(Sender: TObject; AWorksheet: TsWorksheet;
      ARow, ACol: Cardinal; const ASearchText, AReplaceText: String;
      var AConfirmReplacement: TsConfirmReplacementResult);
    procedure PopulateOptions;
  public
    { public declarations }
    procedure Execute(AWorkbook: TsWorkbook);
    property Workbook: TsWorkbook read FWorkbook;
    property SearchParams: TsSearchParams read GetSearchParams write SetSearchParams;
    property ReplaceParams: TsReplaceParams read GetReplaceParams write SetReplaceParams;
    property OnFound: TsSearchEvent read FOnFound write FOnFound;
  end;

var
  SearchForm: TSearchForm;

  DefaultSearchParams: TsSearchParams = (
    SearchText: '';
    Options: [];
    Within: swWorksheet
  );
  DefaultReplaceParams: TsReplaceParams = (
    ReplaceText: '';
    Options: [roConfirm]
  );


implementation

{$R *.lfm}

uses
  fpsUtils;

const
  MAX_SEARCH_ITEMS      = 10;

  // Search & replace
  COMPARE_ENTIRE_CELL   = 0;
  MATCH_CASE            = 1;
  REGULAR_EXPRESSION    = 2;
  SEARCH_ALONG_ROWS     = 3;
  CONTINUE_AT_START_END = 4;
  // Replace only
  REPLACE_ENTIRE_CELL   = 5;
  REPLACE_ALL           = 6;
  CONFIRM_REPLACEMENT   = 7;

  BASE_HEIGHT           = 340;  // Design height of SearchForm

  SEARCH_TAB            = 0;
  REPLACE_TAB           = 1;

var
  CONFIRM_REPLACEMENT_DLG_X: Integer = -1;
  CONFIRM_REPLACEMENT_DLG_Y: Integer = -1;

{ TSearchForms }

procedure TSearchForm.ConfirmReplacementHandler(Sender: TObject;
  AWorksheet: TsWorksheet; ARow, ACol: Cardinal; const ASearchText, AReplaceText: String;
  var AConfirmReplacement: TsConfirmReplacementResult);
var
  F: TForm;
begin
  Unused(AWorksheet, ARow, ACol);
  Unused(ASearchText, AReplaceText);
  F := CreateMessageDialog('Replace?', mtConfirmation, [mbYes, mbNo, mbCancel]);
  try
    if (CONFIRM_REPLACEMENT_DLG_X = -1) then
      F.Position := poMainformCenter
    else begin
      F.Position := poDesigned;
      F.Left := CONFIRM_REPLACEMENT_DLG_X;
      F.Top := CONFIRM_REPLACEMENT_DLG_Y;
    end;
    case F.ShowModal of
      mrYes: AConfirmReplacement := crReplace;
      mrNo : AConfirmReplacement := crIgnore;
      mrCancel: AConfirmReplacement := crAbort;
    end;
    CONFIRM_REPLACEMENT_DLG_X := F.Left;
    CONFIRM_REPLACEMENT_DLG_Y := F.Top;
  finally
    F.Free;
  end;
  {
  case MessageDlg('Replace?', mtConfirmation, [mbYes, mbNo, mbCancel], 0) of
    mrYes: AConfirmReplacement := crReplace;
    mrNo : AConfirmReplacement := crIgnore;
    mrCancel: AConfirmReplacement := crAbort;
  end;
  }
end;

procedure TSearchForm.Execute(AWorkbook: TsWorkbook);
begin
  FWorkbook := AWorkbook;
  Show;
end;

procedure TSearchForm.ExecuteClick(Sender: TObject);
var
  sp: TsSearchParams;
  rp: TsReplaceParams;
  found: Boolean;
  crs: TCursor;
begin
  sp := GetSearchParams;
  if sp.SearchText = '' then
    exit;

  if TabControl.TabIndex = REPLACE_TAB then
    rp := GetReplaceParams;

  if CbSearchText.Items.IndexOf(sp.SearchText) = -1 then
  begin
    CbSearchText.Items.Insert(0, sp.SearchText);
    while CbSearchText.Items.Count > MAX_SEARCH_ITEMS do
      CbSearchText.Items.Delete(CbSearchText.Items.Count-1);
  end;

  if (TabControl.TabIndex = REPLACE_TAB) and
     (CbReplaceText.Items.IndexOf(rp.ReplaceText) = -1) then
  begin
    CbReplaceText.items.Insert(0, rp.ReplaceText);
    while CbReplaceText.Items.Count > MAX_SEARCH_ITEMS do
      CbReplaceText.Items.Delete(CbReplaceText.Items.Count-1);
  end;

  crs := Screen.Cursor;
  try
    Screen.Cursor := crHourglass;
    if FSearchEngine = nil then
    begin
      FSearchEngine := TsSearchEngine.Create(FWorkbook);
      FSearchEngine.OnConfirmReplacement := @ConfirmReplacementHandler;
      if (soBackward in sp.Options) then
        Include(sp.Options, soBackward) else
        Exclude(sp.Options, soBackward);
      case Tabcontrol.TabIndex of
        0: found := FSearchEngine.FindFirst(sp, FFoundWorksheet, FFoundRow, FFoundCol);
        1: found := FSearchEngine.ReplaceFirst(sp, rp, FFoundWorksheet, FFoundRow, FFoundCol);
      end;
    end else
    begin
      // Adjust "backward" option according to the button clicked
      if (Sender = BtnSearchBack) then
        Include(sp.Options, soBackward) else
        Exclude(sp.Options, soBackward);
      // Begin searching at current position
      Exclude(sp.Options, soEntireDocument);
      // User may select a different worksheet/different cell to continue search!
      FFoundWorksheet := FWorkbook.ActiveWorksheet;
      FFoundRow := FFoundWorksheet.ActiveCellRow;
      FFoundCol := FFoundWorksheet.ActiveCellCol;
      case TabControl.TabIndex of
        0: found := FSearchEngine.FindFirst(sp, FFoundWorksheet, FFoundRow, FFoundCol);
        1: found := FSearchEngine.ReplaceFirst(sp, rp, FFoundWorksheet, FFoundRow, FFoundCol);
      end;
    end;

  finally
    Screen.Cursor := crs;
  end;

  if Assigned(FOnFound) then
    FOnFound(self, found, FFoundWorksheet, FFoundRow, FFoundCol);

  BtnSearchBack.Visible := true;
  BtnSearch.Caption := 'Next';
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
 {$IFDEF MSWINDOWS}
  SearchTextPanel.Color := clNone;
  ReplaceTextPanel.Color := clNone;
  SearchParamsPanel.Color := clNone;
 {$ENDIF}
  Position := poMainFormCenter;
  PopulateOptions;
end;

procedure TSearchForm.FormShow(Sender: TObject);
begin
  BtnSearch.Caption := 'Search';
  BtnSearchBack.Visible := false;

  FFoundCol := UNASSIGNED_ROW_COL_INDEX;
  FFoundRow := UNASSIGNED_ROW_COL_INDEX;
  FFoundWorksheet := nil;
end;

function TSearchForm.GetReplaceParams: TsReplaceParams;
begin
  if TabControl.TabIndex = 0 then
    Result := FReplaceParams
  else
  begin
    Result.ReplaceText := CbReplaceText.Text;
    Result.Options := [];
    if CgOptions.Checked[REPLACE_ENTIRE_CELL] then
      Include(Result.Options, roReplaceEntireCell);
    if CgOptions.Checked[REPLACE_ALL] then
      Include(Result.Options, roReplaceAll);
    if CgOptions.Checked[CONFIRM_REPLACEMENT] then
      Include(Result.Options, roConfirm);
    FReplaceParams := Result;
  end;
end;

function TSearchForm.GetSearchParams: TsSearchParams;
begin
  Result.SearchText := CbSearchText.Text;
  Result.Options := [];
  if CgOptions.Checked[COMPARE_ENTIRE_CELL] then
    Include(Result.Options, soCompareEntireCell);
  if CgOptions.Checked[MATCH_CASE] then
    Include(Result.Options, soMatchCase);
  if CgOptions.Checked[REGULAR_EXPRESSION] then
    Include(Result.Options, soRegularExpr);
  if CgOptions.Checked[SEARCH_ALONG_ROWS] then
    Include(Result.Options, soAlongRows);
  if CgOptions.Checked[CONTINUE_AT_START_END] then
    Include(Result.Options, soWrapDocument);
  if RgSearchStart.ItemIndex = 1 then
    Include(Result.Options, soEntireDocument);
  Result.Within := TsSearchWithin(RgSearchWithin.ItemIndex);
end;

procedure TSearchForm.PopulateOptions;
begin
  with CgOptions.Items do
  begin
    Clear;
    Add('Compare entire cell');
    Add('Match case');
    Add('Regular expression');
    Add('Search along rows');
    Add('Continue at start/end');
    if TabControl.TabIndex = REPLACE_TAB then
    begin
      Add('Replace entire cell');
      Add('Replace all');
      Add('Confirm replacement');
    end;
  end;
end;

procedure TSearchForm.SetSearchParams(const AValue: TsSearchParams);
begin
  CbSearchText.Text := Avalue.SearchText;
  CgOptions.Checked[COMPARE_ENTIRE_CELL] := (soCompareEntireCell in AValue.Options);
  CgOptions.Checked[MATCH_CASE] := (soMatchCase in AValue.Options);
  CgOptions.Checked[REGULAR_EXPRESSION] := (soRegularExpr in Avalue.Options);
  CgOptions.Checked[SEARCH_ALONG_ROWS] := (soAlongRows in AValue.Options);
  CgOptions.Checked[CONTINUE_AT_START_END] := (soWrapDocument in Avalue.Options);
  RgSearchWithin.ItemIndex := ord(AValue.Within);
  RgSearchStart.ItemIndex := ord(soEntireDocument in AValue.Options);
end;

procedure TSearchForm.SetReplaceParams(const AValue: TsReplaceParams);
begin
  FReplaceParams := AValue;
  if TabControl.TabIndex = REPLACE_TAB then
  begin
    CbReplaceText.Text := AValue.ReplaceText;
    CgOptions.Checked[REPLACE_ENTIRE_CELL] := (roReplaceEntireCell in AValue.Options);
    CgOptions.Checked[REPLACE_ALL] := (roReplaceAll in AValue.Options);
    CgOptions.Checked[CONFIRM_REPLACEMENT] := (roConfirm in AValue.Options);
  end;
end;

procedure TSearchForm.TabControlChange(Sender: TObject);
var
  h, d: Integer;
begin
  ReplaceTextPanel.Visible := (TabControl.TabIndex = REPLACE_TAB);
  PopulateOptions;
  SetSearchParams(FSearchParams);
  SetReplaceParams(FReplaceParams);
  h := RgSearchStart.Top + RgSearchStart.Height - CgOptions.Top;
  if TabControl.TabIndex = 0 then
  begin
    CgOptions.Height := h;
    Height := BASE_HEIGHT - ReplaceTextPanel.Height;
  end else
  begin
    d := 3 * 16;
    CgOptions.Height := h + d;
    Height := BASE_HEIGHT + d;
  end;
end;

procedure TSearchForm.TabControlChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
  AllowChange := true;
  FSearchParams := GetSearchParams;
  FReplaceParams := GetReplaceParams;
end;


end.

