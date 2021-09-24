unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, Forms, Controls, Graphics, Dialogs, DBGrids, DBCtrls,
  ExtCtrls, Menus, StdCtrls, Buttons, ActnList, ComCtrls, fpsDataset, Types;

type

  { TMainForm }

  TMainForm = class(TForm)
    acFindCity: TAction;
    acFilterByCountry: TAction;
    acNoFilter: TAction;
    acSetBookmark: TAction;
    acGotoBookmark: TAction;
    acClearBookmark: TAction;
    acSortAsc: TAction;
    acSortDesc: TAction;
    ActionList: TActionList;
    cmbFilter: TComboBox;
    DataSource: TDataSource;
    DBGrid: TDBGrid;
    Dataset: TsWorksheetDataset;
    DBNavigator1: TDBNavigator;
    ImageList16: TImageList;
    ImageList12: TImageList;
    Label1: TLabel;
    MenuItem1: TMenuItem;
    mnuClearBookmarkParent: TMenuItem;
    MenuItem5: TMenuItem;
    mnuSetBookmark: TMenuItem;
    mnuGotoBookmark: TMenuItem;
    mnuClearBookmark: TMenuItem;
    mnuFindCity: TMenuItem;
    mnuClearBookmark1: TMenuItem;
    mnuClearBookmark2: TMenuItem;
    mnuClearBookmark3: TMenuItem;
    mnuSetBookmark3: TMenuItem;
    mnuSetBookmark2: TMenuItem;
    mnuSetBookmark1: TMenuItem;
    mnuGotoBookmark3: TMenuItem;
    mnuGotoBookmark2: TMenuItem;
    mnuGotoBookmark1: TMenuItem;
    mnuSetBookmarkParent: TMenuItem;
    MenuItem3: TMenuItem;
    mnuGotoBookmarkParent: TMenuItem;
    N1: TMenuItem;
    mnuNoFilter: TMenuItem;
    mnuFilterByCountry: TMenuItem;
    mnuSortASC: TMenuItem;
    mnuSortDESC: TMenuItem;
    DBGridPopupMenu: TPopupMenu;
    Panel1: TPanel;
    Panel2: TPanel;
    BookmarkDropdown: TPopupMenu;
    ToolBar: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    btnBookmark1: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    btnBookmark2: TToolButton;
    btnBookmark3: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    procedure acClearBookmarkExecute(Sender: TObject);
    procedure acFilterByCountryExecute(Sender: TObject);
    procedure acFindCityExecute(Sender: TObject);
    procedure acGotoBookmarkExecute(Sender: TObject);
    procedure acNoFilterExecute(Sender: TObject);
    procedure acSetBookmarkExecute(Sender: TObject);
    procedure acSortAscExecute(Sender: TObject);
    procedure acSortDescExecute(Sender: TObject);
    procedure BookmarkDropdownPopup(Sender: TObject);
    procedure btnBookmark1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure cmbFilterEditingDone(Sender: TObject);
    procedure DatasetAfterOpen(ADataSet: TDataSet);
    procedure DBGridTitleClick(Column: TColumn);
    procedure FormCreate(Sender: TObject);
    procedure DBGridPopupMenuPopup(Sender: TObject);
    procedure mnuFindCityClick(Sender: TObject);
  private
    FBookmarks: array[1..3] of TBookmark;
    FSortColumn: TColumn;
    procedure GetUniqueFieldValues(AField: TField; AList: TStrings);
    procedure mnuFilterHandler(Sender: TObject);
    procedure PrepareFilter;

  public

  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

uses
  StrUtils, fpsTypes;


{ TMainForm }

procedure TMainForm.acFilterByCountryExecute(Sender: TObject);
begin
  Dataset.Filtered := false;
  if (cmbFilter.Text = '') or (cmbFilter.Text = 'all countries') then
    cmbFilter.ItemIndex := 0
  else
  begin
    Dataset.Filter := 'Country = "' + cmbFilter.Text + '"';
    Dataset.Filtered := true;
  end;
end;

procedure TMainForm.acClearBookmarkExecute(Sender: TObject);
var
  ac: TAction;
  mnu: TComponent;
  idx: Integer;
begin
  if not (Sender is TAction) then
    exit;
  mnu := TAction(Sender).ActionComponent;
  idx := mnu.Tag;
  if idx <= 0 then
    exit;
  if Dataset.BookmarkValid(FBookmarks[idx]) then
  begin
    Dataset.FreeBookmark(FBookmarks[idx]);
    FBookmarks[idx] := nil;
  end;
end;

procedure TMainForm.acFindCityExecute(Sender: TObject);
var
  s: String;
begin
  s := InputBox('Find city', 'City', '');
  if s <> '' then
  begin
    if not Dataset.Locate('City', s, [loCaseInsensitive, loPartialKey]) then
      ShowMessage('Not found.');
  end;
end;

procedure TMainForm.acGotoBookmarkExecute(Sender: TObject);
var
  ac: TAction;
  mnu: TComponent;
  idx: Integer;
begin
  if not (Sender is TAction) then
    exit;
  mnu := TAction(Sender).ActionComponent;
  idx := mnu.Tag;
  if idx <= 0 then
    exit;
  if Dataset.BookmarkValid(FBookmarks[idx]) then
    try
      Dataset.GotoBookmark(FBookmarks[idx]);
    except
      MessageDlg('Bookmark not found (filtered?)', mtError, [mbOK], 0);
    end;
end;

procedure TMainForm.acNoFilterExecute(Sender: TObject);
begin
  cmbFilter.ItemIndex := 0;

  Dataset.Filtered := false;
  Dataset.Filter := '';
end;

procedure TMainForm.acSetBookmarkExecute(Sender: TObject);
var
  ac: TAction;
  mnu: TComponent;
  idx: Integer;
begin
  if not (Sender is TAction) then
    exit;
  mnu := TAction(Sender).ActionComponent;
  idx := mnu.Tag;
  if idx <= 0 then
    exit;
  FBookmarks[idx] := Dataset.Getbookmark;
end;

procedure TMainForm.acSortAscExecute(Sender: TObject);
begin
  if FSortColumn <> nil then FSortColumn.Title.ImageIndex := -1;
  FSortColumn := DBGrid.SelectedColumn;
  FSortColumn.Title.ImageIndex := 0;
  Dataset.SortOnField(FSortColumn.FieldName);
end;

procedure TMainForm.acSortDescExecute(Sender: TObject);
begin
  if FSortColumn <> nil then FSortColumn.Title.ImageIndex := -1;
  FSortColumn := DBGrid.SelectedColumn;
  FSortColumn.Title.ImageIndex := 1;
  Dataset.SortOnField(DBGrid.SelectedColumn.FieldName, [ssoDescending]);
end;

procedure TMainForm.BookmarkDropdownPopup(Sender: TObject);
var
  idx: Integer;
  dropdown: TPopupMenu;
  btn: TToolButton;
begin
  if not (Sender is TPopupMenu) then
    exit;

  dropdown := TPopupMenu(Sender);
  idx := dropdown.Tag;
  mnuSetBookmark.Tag := idx;
  mnuGotoBookmark.Tag := idx;
  mnuClearBookmark.Tag := idx;

  acSetBookmark.Enabled := true;
  acGotoBookmark.Enabled := Dataset.BookmarkValid(FBookmarks[idx]);
  acClearBookmark.Enabled := Dataset.BookmarkValid(FBookmarks[idx]);
end;

procedure TMainForm.btnBookmark1MouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  with TToolButton(Sender) do
    DropDownMenu.Tag := Tag;
end;

procedure TMainForm.cmbFilterEditingDone(Sender: TObject);
begin
  acFilterByCountryExecute(nil);
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
  i: Integer;
begin
  // Open spreadsheet file as dataset
  Dataset.FileName := 'Temperatures.xlsx';
  Dataset.Open;

  // Tailor the columns of the DBGrid
  DBGrid.Columns[0].Width := 100;
  DBGrid.Columns[1].Width := 100;
  for i := 2 to DBGrid.Columns.Count-1 do
    with DBGrid.Columns[i] do begin
      Width := 64;
      DisplayFormat := '0.0'; // Avoid too many decimal places in floating point fields.
      Title.Alignment := taCenter;
    end;

  // Prepare bookmarks
  FBookmarks[1] := nil;
  FBookmarks[2] := nil;
  FBookmarks[3] := nil;

  // Narrower input box
  cInputQueryEditSizePercents := 0;
end;

procedure TMainForm.DBGridPopupMenuPopup(Sender: TObject);
begin
  mnuGotoBookmark1.Enabled := Dataset.BookmarkValid(FBookmarks[1]);
  mnuGotoBookmark2.Enabled := Dataset.BookmarkValid(FBookmarks[2]);
  mnuGotoBookmark3.Enabled := Dataset.BookmarkValid(FBookmarks[3]);

  mnuClearBookmark1.Enabled := Dataset.BookmarkValid(FBookmarks[1]);
  mnuClearBookmark2.Enabled := Dataset.BookmarkValid(FBookmarks[2]);
  mnuClearBookmark3.Enabled := Dataset.BookmarkValid(FBookmarks[3]);
end;

{ Sorts the grid (and worksheet) when a grid header is clicked. A sort indicator
  image is displayed at the right of the column title. Requires an ImageList
  assigned to the grid's TitleImageList having the image for ascending and
  descending sorts at index 0 and 1, respectively. }
procedure TMainForm.DBGridTitleClick(Column: TColumn);
var
  options: TsSortOptions;
begin
  options := [];  // [] --> ascending sort

  if FSortColumn = Column then
  // Previously selected sort column was clicked another time...
  begin
    // Toggle between ascending and descending sort images
    FSortColumn.Title.ImageIndex := (FSortColumn.Title.ImageIndex + 1) mod 2;
    if FSortColumn.Title.ImageIndex = 1 then
      options := [ssoDescending];
  end
  else
  // A previously unsorted column was clicked...
  begin
    // Remove sort image from old sort column
    if FSortColumn <> nil then FSortColumn.Title.ImageIndex := -1;
    // Store clicked column as new SortColumn
    FSortColumn := Column;
    // Set sort image index to "ascending sort"
    FSortColumn.Title.ImageIndex := 0;
  end;

  // Execute the sorting operation.
  Dataset.SortOnField(FSortColumn.Field.FieldName, options);
end;

procedure TMainForm.DatasetAfterOpen(ADataSet: TDataSet);
begin
  PrepareFilter;
end;

procedure TMainForm.GetUniqueFieldValues(AField: TField; AList: TStrings);
var
  bm: TBookmark;
  L: TStringList;
begin
  bm := Dataset.GetBookmark;
  Dataset.DisableControls;
  L := TStringList.Create;
  try
    L.Sorted := true;
    L.Duplicates := dupIgnore;
    Dataset.First;
    while not Dataset.EOF do
    begin
      L.Add(AField.AsString);
      Dataset.Next;
    end;
    AList.Assign(L);
  finally
    L.Free;
    if Dataset.BookmarkValid(bm) then
    begin
      Dataset.GotoBookmark(bm);
      Dataset.FreeBookmark(bm);
    end;
    Dataset.EnableControls;
  end;
end;

procedure TMainForm.mnuFindCityClick(Sender: TObject);
var
  s: String;
begin
  s := InputBox('Find city', 'City', '');
  if s <> '' then
  begin
    if not Dataset.Locate('City', s, [loCaseInsensitive, loPartialKey]) then
      ShowMessage('Not found.');
  end;
end;

procedure TMainForm.mnuFilterHandler(Sender: TObject);
begin
  cmbFilter.Text := (Sender as TMenuItem).Caption;

  Dataset.Filtered := false;
  Dataset.Filter := 'Country = "' + (Sender as TMenuItem).Caption + '"';
  Dataset.Filtered := true;
end;

procedure TMainForm.PrepareFilter;
var
  L: TStrings;
  mnu: TMenuItem;
  i: Integer;
begin
  L := TStringList.Create;
  try
    GetUniqueFieldValues(DBGrid.SelectedColumn.Field, L);
    for i := 0 to L.Count-1 do begin
      mnu := TMenuItem.Create(mnuFilterByCountry);
      mnu.Caption := L[i];
      mnu.OnClick := @mnuFilterHandler;
      MnuFilterByCountry.Add(mnu);
    end;

    cmbFilter.Items.Assign(L);
    cmbFilter.Items.Insert(0, 'all countries');
    cmbFilter.ItemIndex := 0;

  finally
    L.Free;
  end;
end;

end.

