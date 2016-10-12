unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, ActnList, Menus, StdActns, Buttons,
  fpstypes, fpspreadsheet, fpspreadsheetctrls, fpspreadsheetgrid, fpsActions,
  fpsRegFileFormats, fpsSYLK, xlsxml, Grids, Types;

type

  { TMainForm }

  TMainForm = class(TForm)
    AcRowDelete: TAction;
    AcColDelete: TAction;
    AcRowAdd: TAction;
    AcColAdd: TAction;
    AcSettingsCSVParams: TAction;
    AcSettingsCurrency: TAction;
    AcSettingsFormatSettings: TAction;
    AcSearch: TAction;
    AcShowGridLines: TAction;
    AcShowHeaders: TAction;
    AcFrozenRows: TAction;
    AcFrozenCols: TAction;
    AcAutoRowHeights: TAction;
    AcWorksheetRTL: TAction;
    AcViewInspector: TAction;
    ActionList: TActionList;
    AcFileExit: TFileExit;
    AcFileOpen: TFileOpen;
    AcFileSaveAs: TFileSaveAs;
    ImageList: TImageList;
    MainMenu: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem100: TMenuItem;
    MenuItem101: TMenuItem;
    MenuItem102: TMenuItem;
    MenuItem103: TMenuItem;
    MenuItem104: TMenuItem;
    MenuItem105: TMenuItem;
    MenuItem106: TMenuItem;
    MenuItem107: TMenuItem;
    MenuItem108: TMenuItem;
    MenuItem109: TMenuItem;
    MenuItem110: TMenuItem;
    MenuItem111: TMenuItem;
    MenuItem112: TMenuItem;
    MenuItem113: TMenuItem;
    MenuItem114: TMenuItem;
    MenuItem115: TMenuItem;
    MenuItem116: TMenuItem;
    MenuItem117: TMenuItem;
    MenuItem118: TMenuItem;
    MenuItem119: TMenuItem;
    MenuItem120: TMenuItem;
    MenuItem121: TMenuItem;
    MenuItem122: TMenuItem;
    MenuItem123: TMenuItem;
    MenuItem124: TMenuItem;
    MenuItem125: TMenuItem;
    MenuItem126: TMenuItem;
    MenuItem127: TMenuItem;
    MenuItem128: TMenuItem;
    MenuItem129: TMenuItem;
    MenuItem130: TMenuItem;
    MenuItem131: TMenuItem;
    MenuItem132: TMenuItem;
    MenuItem133: TMenuItem;
    MenuItem134: TMenuItem;
    MenuItem135: TMenuItem;
    MenuItem136: TMenuItem;
    MenuItem137: TMenuItem;
    MenuItem138: TMenuItem;
    MenuItem139: TMenuItem;
    MenuItem140: TMenuItem;
    MenuItem141: TMenuItem;
    MenuItem142: TMenuItem;
    MenuItem143: TMenuItem;
    MenuItem144: TMenuItem;
    MenuItem145: TMenuItem;
    MenuItem146: TMenuItem;
    MenuItem148: TMenuItem;
    MenuItem149: TMenuItem;
    MenuItem150: TMenuItem;
    MenuItem151: TMenuItem;
    MenuItem152: TMenuItem;
    MenuItem153: TMenuItem;
    MenuItem154: TMenuItem;
    MenuItem155: TMenuItem;
    MenuItem156: TMenuItem;
    MenuItem157: TMenuItem;
    MenuItem158: TMenuItem;
    MenuItem159: TMenuItem;
    MnuZoom: TMenuItem;
    MenuItem147: TMenuItem;
    MnuSettings: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem34: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem36: TMenuItem;
    MenuItem37: TMenuItem;
    MenuItem38: TMenuItem;
    MenuItem39: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem40: TMenuItem;
    MenuItem41: TMenuItem;
    MenuItem42: TMenuItem;
    MenuItem43: TMenuItem;
    MenuItem44: TMenuItem;
    MenuItem45: TMenuItem;
    MenuItem46: TMenuItem;
    MenuItem47: TMenuItem;
    MenuItem48: TMenuItem;
    MenuItem49: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem50: TMenuItem;
    MenuItem51: TMenuItem;
    MenuItem52: TMenuItem;
    MenuItem53: TMenuItem;
    MenuItem54: TMenuItem;
    MenuItem55: TMenuItem;
    MenuItem56: TMenuItem;
    MenuItem57: TMenuItem;
    MenuItem58: TMenuItem;
    MenuItem59: TMenuItem;
    MenuItem60: TMenuItem;
    MenuItem61: TMenuItem;
    MenuItem62: TMenuItem;
    MenuItem63: TMenuItem;
    MenuItem64: TMenuItem;
    MenuItem65: TMenuItem;
    MenuItem66: TMenuItem;
    MenuItem67: TMenuItem;
    MenuItem68: TMenuItem;
    MenuItem69: TMenuItem;
    MenuItem70: TMenuItem;
    MenuItem71: TMenuItem;
    MenuItem72: TMenuItem;
    MenuItem73: TMenuItem;
    MenuItem74: TMenuItem;
    MenuItem75: TMenuItem;
    MenuItem76: TMenuItem;
    MenuItem77: TMenuItem;
    MenuItem78: TMenuItem;
    MenuItem79: TMenuItem;
    MenuItem80: TMenuItem;
    MenuItem81: TMenuItem;
    MenuItem82: TMenuItem;
    MenuItem83: TMenuItem;
    MenuItem84: TMenuItem;
    MenuItem85: TMenuItem;
    MenuItem86: TMenuItem;
    MenuItem87: TMenuItem;
    MenuItem88: TMenuItem;
    MenuItem89: TMenuItem;
    MenuItem9: TMenuItem;
    MenuItem90: TMenuItem;
    MenuItem91: TMenuItem;
    MenuItem95: TMenuItem;
    MenuItem96: TMenuItem;
    MenuItem97: TMenuItem;
    MenuItem98: TMenuItem;
    MenuItem99: TMenuItem;
    MnuColumn: TMenuItem;
    MenuItem93: TMenuItem;
    MenuItem94: TMenuItem;
    MnuAddWorksheet: TMenuItem;
    MnuRow: TMenuItem;
    MenuItem92: TMenuItem;
    MnuView: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MnuFormat: TMenuItem;
    MnuFile: TMenuItem;
    MnuWorksheet: TMenuItem;
    MnuAddSheet: TMenuItem;
    MnuEdit: TMenuItem;
    OpenDialog: TOpenDialog;
    OpenDialog1: TOpenDialog;
    CellEdit: TsCellEdit;
    CellIndicator: TsCellIndicator;
    AcFontBold: TsFontStyleAction;
    AcFontItalic: TsFontStyleAction;
    AcVertAlignTop: TsVertAlignmentAction;
    AcVertAlignCenter: TsVertAlignmentAction;
    AcVertAlignBottom: TsVertAlignmentAction;
    AcHorAlignLeft: TsHorAlignmentAction;
    AcHorAlignCenter: TsHorAlignmentAction;
    AcHorAlignRight: TsHorAlignmentAction;
    AcTextRotHor: TsTextRotationAction;
    AcTextRot90CW: TsTextRotationAction;
    AcTextRot90CCW: TsTextRotationAction;
    AcTextRotStacked: TsTextRotationAction;
    AcWordWrap: TsWordwrapAction;
    AcNumFormatFixed: TsNumberFormatAction;
    AcNumFormatFixedTh: TsNumberFormatAction;
    AcNumFormatPercentage: TsNumberFormatAction;
    AcNumFormatCurrency: TsNumberFormatAction;
    AcNumFormatCurrencyRed: TsNumberFormatAction;
    Panel2: TPanel;
    PuPaste: TPopupMenu;
    PuBorders: TPopupMenu;
    PuTimeFormat: TPopupMenu;
    PuDateFormat: TPopupMenu;
    PuCurrencyFormat: TPopupMenu;
    PuNumFormat: TPopupMenu;
    AcNumFormatGeneral: TsNumberFormatAction;
    AcNumFormatExp: TsNumberFormatAction;
    AcNumFormatDateTime: TsNumberFormatAction;
    AcNumFormatLongDate: TsNumberFormatAction;
    AcNumFormatShortDate: TsNumberFormatAction;
    AcNumFormatLongTime: TsNumberFormatAction;
    AcNumFormatShortTime: TsNumberFormatAction;
    AcNumFormatLongTimeAM: TsNumberFormatAction;
    AcNumFormatShortTimeAM: TsNumberFormatAction;
    AcNumFormatTimeInterval: TsNumberFormatAction;
    AcIncDecimals: TsDecimalsAction;
    AcDecDecimals: TsDecimalsAction;
    AcCellFontDialog: TsFontDialogAction;
    AcBackgroundColorDialog: TsBackgroundColorDialogAction;
    AcCellBorderTop: TsCellBorderAction;
    AcCellBorderBottom: TsCellBorderAction;
    AcCellBorderLeft: TsCellBorderAction;
    AcCellBorderRight: TsCellBorderAction;
    AcCellBorderInnerHor: TsCellBorderAction;
    AcCellBorderInnerVert: TsCellBorderAction;
    AcCellBorderAllHor: TsCellBorderAction;
    AcCellBorderBottomThick: TsCellBorderAction;
    AcCellBorderBottomDbl: TsCellBorderAction;
    AcCellBorderAllOuter: TsCellBorderAction;
    AcCellBorderNone: TsNoCellBordersAction;
    AcCellBorderAllOuterThick: TsCellBorderAction;
    AcCellBorderTopBottomThick: TsCellBorderAction;
    AcCellBorderTopBottomDbl: TsCellBorderAction;
    AcCellBorderAll: TsCellBorderAction;
    AcCellBorderAllVert: TsCellBorderAction;
    AcCopyFormat: TsCopyAction;
    FontColorCombobox: TsCellCombobox;
    BackgroundColorCombobox: TsCellCombobox;
    FontnameCombo: TsCellCombobox;
    FontsizeCombo: TsCellCombobox;
    AcMergeCells: TsMergeAction;
    AcCopyToClipboard: TsCopyAction;
    AcCutToClipboard: TsCopyAction;
    AcPasteAllFromClipboard: TsCopyAction;
    AcPasteValueFromClipboard: TsCopyAction;
    AcPasteFormatFromClipboard: TsCopyAction;
    AcPasteFormulaFromClipboard: TsCopyAction;
    AcCommentNew: TsCellCommentAction;
    AcCommentEdit: TsCellCommentAction;
    AcCommentDelete: TsCellCommentAction;
    AcHyperlinkNew: TsCellHyperlinkAction;
    AcHyperlinkEdit: TsCellHyperlinkAction;
    AcHyperlinkDelete: TsCellHyperlinkAction;
    AcNumFormatFraction2: TsNumberFormatAction;
    AcNumFormatFraction1: TsNumberFormatAction;
    AcNumFormatFraction3: TsNumberFormatAction;
    AcNumFormatDayMonth: TsNumberFormatAction;
    AcNumFormatMonthYear: TsNumberFormatAction;
    AcNumFormatCustom: TsNumberFormatAction;
    AcCellBorderDiagUp: TsCellBorderAction;
    AcCellBorderDiagDown: TsCellBorderAction;
    AcNumFormatText: TsNumberFormatAction;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    AcZoom100: TsWorksheetZoomAction;
    AcZoom90: TsWorksheetZoomAction;
    AcZoom30: TsWorksheetZoomAction;
    AcZoom50: TsWorksheetZoomAction;
    AcZoom75: TsWorksheetZoomAction;
    AcZoom150: TsWorksheetZoomAction;
    AcZoom200: TsWorksheetZoomAction;
    AcZoom300: TsWorksheetZoomAction;
    AcZoom500: TsWorksheetZoomAction;
    ToolBar2: TToolBar;
    ToolBar3: TToolBar;
    ToolButton1: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    AcFontUnderline: TsFontStyleAction;
    AcFontStrikeout: TsFontStyleAction;
    InspectorSplitter: TSplitter;
    Inspector: TsSpreadsheetInspector;
    InspectorTabControl: TTabControl;
    AcAddWorksheet: TsWorksheetAddAction;
    AcDeleteWorksheet: TsWorksheetDeleteAction;
    acRenameWorksheet: TsWorksheetRenameAction;
    ToolBar1: TToolBar;
    ToolButton10: TToolButton;
    ToolButton2: TToolButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    ToolButton26: TToolButton;
    ToolButton27: TToolButton;
    ToolButton28: TToolButton;
    ToolButton29: TToolButton;
    ToolButton3: TToolButton;
    ToolButton30: TToolButton;
    ToolButton31: TToolButton;
    TbBorders: TToolButton;
    ToolButton32: TToolButton;
    ToolButton33: TToolButton;
    ToolButton34: TToolButton;
    ToolButton35: TToolButton;
    ToolButton36: TToolButton;
    ToolButton37: TToolButton;
    ToolButton38: TToolButton;
    ToolButton39: TToolButton;
    TbCommentAdd: TToolButton;
    ToolButton4: TToolButton;
    ToolButton40: TToolButton;
    ToolButton41: TToolButton;
    ToolButton42: TToolButton;
    ToolButton43: TToolButton;
    ToolButton44: TToolButton;
    ToolButton45: TToolButton;
    ToolButton46: TToolButton;
    ToolButton47: TToolButton;
    ToolButton48: TToolButton;
    ToolButton49: TToolButton;
    ToolButton5: TToolButton;
    TbCommentDelete: TToolButton;
    TbCommentEdit: TToolButton;
    ToolButton50: TToolButton;
    ToolButton51: TToolButton;
    ToolButton52: TToolButton;
    ToolButton53: TToolButton;
    ToolButton54: TToolButton;
    ToolButton55: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    WorkbookSource: TsWorkbookSource;
    WorkbookTabControl: TsWorkbookTabControl;
    WorksheetGrid: TsWorksheetGrid;
    procedure AcAutoRowHeightsExecute(Sender: TObject);
    procedure AcColAddExecute(Sender: TObject);
    procedure AcColDeleteExecute(Sender: TObject);
    procedure AcFileOpenAccept(Sender: TObject);
    procedure AcFileSaveAsAccept(Sender: TObject);
    procedure AcFileSaveAsBeforeExecute(Sender: TObject);
    procedure AcFrozenColsExecute(Sender: TObject);
    procedure AcFrozenColsUpdate(Sender: TObject);
    procedure AcFrozenRowsExecute(Sender: TObject);
    procedure AcFrozenRowsUpdate(Sender: TObject);
    procedure AcNumFormatCustomGetNumberFormatString(Sender: TObject;
      AWorkbook: TsWorkbook; var ANumFormatStr: String);
    procedure AcRowAddExecute(Sender: TObject);
    procedure AcRowDeleteExecute(Sender: TObject);
    procedure ActionListUpdate(AAction: TBasicAction; var Handled: Boolean);
    procedure AcWorksheetRTLExecute(Sender: TObject);
    procedure AcWorksheetRTLUpdate(Sender: TObject);
    procedure AcSearchExecute(Sender: TObject);
    procedure AcSettingsCSVParamsExecute(Sender: TObject);
    procedure AcSettingsCurrencyExecute(Sender: TObject);
    procedure AcSettingsFormatSettingsExecute(Sender: TObject);
    procedure AcShowGridLinesExecute(Sender: TObject);
    procedure AcShowGridLinesUpdate(Sender: TObject);
    procedure AcShowHeadersExecute(Sender: TObject);
    procedure AcShowHeadersUpdate(Sender: TObject);
    procedure AcViewInspectorExecute(Sender: TObject);
    procedure EditCut1Execute(Sender: TObject);
    procedure ColorComboboxAddColors(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure HyperlinkHandler(Sender: TObject; ACaption: String;
      var AHyperlink: TsHyperlink);
    procedure InspectorTabControlChange(Sender: TObject);
    procedure TSaveDialogTypeChange(Sender: TObject);
    procedure WorksheetGridClickHyperlink(Sender: TObject;
      const AHyperlink: TsHyperlink);
    procedure WorksheetGridMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
  private
    { private declarations }
    FOpenFormats: TsSpreadFormatIDArray;
    FSaveFormats: TsSpreadFormatIDArray;
    procedure SearchClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure SearchFound(Sender: TObject; AFound: Boolean;
      AWorksheet: TsWorksheet; ARow, ACol: Cardinal);
    procedure UpdateCaption;
  protected
    procedure ReadFromIni;
    procedure WriteToIni;
  public
    { public declarations }
  end;

var
  MainForm: TMainForm;


implementation

{$R *.lfm}

uses
  LCLIntf, inifiles, uriparser,
  fpsUtils, fpsCSV,
  sCSVParamsForm, sCurrencyForm, sFormatSettingsForm, sSortParamsForm,
  sHyperlinkForm, sNumFormatForm, sSearchForm;

var
  SEARCH_DLG_POS: TPoint = (X: -1; Y: -1);

function CreateIni: TCustomIniFile;
begin
  Result := TMemIniFile.Create(GetAppConfigFile(false));
end;


{ TMainForm }

procedure TMainForm.AcAutoRowHeightsExecute(Sender: TObject);
begin
  Screen.Cursor := crHourglass;
  try
    WorksheetGrid.UpdateRowHeights(0, true);
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ Adds a column before the active cell }
procedure TMainForm.AcColAddExecute(Sender: TObject);
begin
  WorksheetGrid.InsertCol(WorksheetGrid.Col);
  WorksheetGrid.Col := WorksheetGrid.Col + 1;
end;

{ Deletes the column with the active cell }
procedure TMainForm.AcColDeleteExecute(Sender: TObject);
var
  c: Integer;
begin
  c := WorksheetGrid.Col;
  WorksheetGrid.DeleteCol(c);
  WorksheetGrid.Col := c;
end;

{ Loads the spreadsheet file selected by the AcFileOpen action }
procedure TMainForm.AcFileOpenAccept(Sender: TObject);
var
  crs: TCursor;
begin
  crs := Screen.Cursor;
  Screen.Cursor := crHourglass;
  try
    WorkbookSource.AutodetectFormat := false;
    case AcFileOpen.Dialog.FilterIndex of
      1: WorkbookSource.AutoDetectFormat := true;      // All spreadsheet files
      2: WorkbookSource.AutoDetectFormat := true;      // All Excel files
      else WorkbookSource.FileFormatID := FOpenFormats[AcFileOpen.Dialog.FilterIndex - 3];
         // -3 because FilterIndex is 1-based and there are 2 add'l items at the top.
    end;
    WorkbookSource.FileName := UTF8ToAnsi(AcFileOpen.Dialog.FileName);  // this loads the file
    UpdateCaption;
  finally
    Screen.Cursor := crs;
  end;
end;

{ Saves the spreadsheet to the file selected by the AcFileSaveAs action }
procedure TMainForm.AcFileSaveAsAccept(Sender: TObject);
var
  fmt: TsSpreadFormatID;
  fmts: TsSpreadFormatIDArray;
  ext: String;
begin
  Screen.Cursor := crHourglass;
  try
    fmt := FSaveFormats[AcFileSaveAs.Dialog.FilterIndex - 1];
    ext := ExtractFileExt(ACFileSaveAs.Dialog.Filename);
    WorkbookSource.SaveToSpreadsheetFile(UTF8ToAnsi(AcFileSaveAs.Dialog.FileName), fmt);
    UpdateCaption;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TMainForm.AcFileSaveAsBeforeExecute(Sender: TObject);
var
  i: Integer;
begin
  if WorkbookSource.FileName = '' then
    exit;

  AcFileSaveAs.Dialog.InitialDir := ExtractFileDir(WorkbookSource.FileName);
  AcFileSaveAs.Dialog.FileName := ExtractFileName(WorkbookSource.FileName);

  // Pre-select the file format according to the input file
  if WorkbookSource.Workbook.FileFormatID = sfidUnknown then
    exit;
  for i:=0 to High(FSaveformats) do
    if FSaveFormats[i] = WorkbookSource.Workbook.FileFormatID then begin
      AcFileSaveAs.Dialog.FilterIndex := i + 1;
      break;
    end;
end;

procedure TMainForm.AcFrozenColsExecute(Sender: TObject);
begin
  WorksheetGrid.FrozenCols := WorksheetGrid.GetWorksheetCol(WorksheetGrid.Col);
end;

procedure TMainForm.AcFrozenColsUpdate(Sender: TObject);
begin
  AcFrozenCols.Checked := WorksheetGrid.FrozenCols > 0;
end;

procedure TMainForm.AcFrozenRowsExecute(Sender: TObject);
begin
  WorksheetGrid.FrozenRows := WorksheetGrid.GetWorksheetRow(WorksheetGrid.Row);
end;

procedure TMainForm.AcFrozenRowsUpdate(Sender: TObject);
begin
  AcFrozenRows.Checked := WorksheetGrid.FrozenRows > 0;
end;

procedure TMainForm.AcNumFormatCustomGetNumberFormatString(Sender: TObject;
  AWorkbook: TsWorkbook; var ANumFormatStr: String);
var
  F: TNumFormatForm;
  sample: Double;
begin
  Unused(AWorkbook);
  F := TNumFormatForm.Create(nil);
  try
    F.Position := poMainFormCenter;
    with WorkbookSource.Worksheet do
      sample := ReadAsNumber(ActiveCellRow, ActiveCellCol);
    F.SetData(ANumFormatStr, WorkbookSource.Workbook, sample);
    if F.ShowModal = mrOK then
      ANumFormatStr := F.NumFormatStr;
  finally
    F.Free;
  end;
end;

{ Adds a row before the active cell }
procedure TMainForm.AcRowAddExecute(Sender: TObject);
begin
  WorksheetGrid.InsertRow(WorksheetGrid.Row);
  WorksheetGrid.Row := WorksheetGrid.Row + 1;
end;

{ Deletes the row with the active cell }
procedure TMainForm.AcRowDeleteExecute(Sender: TObject);
var
  r: Integer;
begin
  r := WorksheetGrid.Row;
  WorksheetGrid.DeleteRow(r);
  WorksheetGrid.Row := r;
end;

procedure TMainForm.AcWorksheetRTLExecute(Sender: TObject);
begin
  if AcWorksheetRTL.Checked then
  begin
    if WorksheetGrid.IsRightToLeft then
      WorksheetGrid.Worksheet.BiDiMode := bdLTR else
      WorksheetGrid.Worksheet.BiDiMode := bdRTL;
  end else
    WorksheetGrid.Worksheet.BiDiMode := bdDefault;
end;

procedure TMainForm.AcWorksheetRTLUpdate(Sender: TObject);
begin
  AcWorksheetRTL.Checked := WorksheetGrid.Worksheet.BiDiMode <> bdDefault;
end;

procedure TMainForm.AcSearchExecute(Sender: TObject);
begin
  if SearchForm = nil then
    SearchForm := TSearchForm.Create(self)
  else
  if not SearchForm.Showing then
  begin
    SearchForm.Position := poDesigned;
    SearchForm.Left := SEARCH_DLG_POS.X;
    SearchForm.Top := SEARCH_DLG_POS.Y;
  end else
    SearchForm.BringToFront;
  SearchForm.OnFound := @SearchFound;
  SearchForm.OnClose := @SearchClose;
  SearchForm.SearchParams := DefaultSearchParams;
  SearchForm.ReplaceParams := DefaultReplaceParams;
  SearchForm.Execute(WorkbookSource.Workbook);
end;

procedure TMainForm.AcSettingsCSVParamsExecute(Sender: TObject);
var
  F: TCSVParamsForm;
begin
  F := TCSVParamsForm.Create(nil);
  try
    F.SetParams(fpscsv.CSVParams);
    if F.ShowModal = mrOK then
      F.GetParams(fpscsv.CSVParams);
  finally
    F.Free;
  end;
end;

procedure TMainForm.AcSettingsCurrencyExecute(Sender: TObject);
var
  F: TCurrencyForm;
begin
  F := TCurrencyForm.Create(nil);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

procedure TMainForm.AcSettingsFormatSettingsExecute(Sender: TObject);
var
  F: TFormatSettingsForm;
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  F := TFormatSettingsForm.Create(nil);
  try
    F.FormatSettings := WorksheetGrid.Workbook.FormatSettings;
    if F.ShowModal = mrOK then
    begin
      WorksheetGrid.Workbook.FormatSettings := F.FormatSettings;
      WorksheetGrid.Invalidate;
    end;
  finally
    F.Free;
  end;
end;

procedure TMainForm.AcShowGridLinesExecute(Sender: TObject);
begin
  WorksheetGrid.ShowGridLines := AcShowGridLines.Checked;
end;

procedure TMainForm.AcShowGridLinesUpdate(Sender: TObject);
begin
  AcShowGridLines.Checked := WorksheetGrid.ShowGridLines;
end;

procedure TMainForm.AcShowHeadersExecute(Sender: TObject);
begin
  WorksheetGrid.ShowHeaders := AcShowHeaders.Checked;
end;

procedure TMainForm.AcShowHeadersUpdate(Sender: TObject);
begin
  AcShowHeaders.Checked := WorksheetGrid.ShowHeaders;
end;

procedure TMainForm.ActionListUpdate(AAction: TBasicAction; var Handled: Boolean
  );
begin
  if AAction = AcAutoRowHeights then
    AcAutoRowHeights.Enabled := WorkbookSource.Worksheet <> nil;
end;

{ Toggles the spreadsheet inspector on and off }
procedure TMainForm.AcViewInspectorExecute(Sender: TObject);
begin
  InspectorTabControl.Visible := AcViewInspector.Checked;
  InspectorSplitter.Visible := AcViewInspector.Checked;
  InspectorSplitter.Left := 0;
  // Make sure that the splitter is always at the left of the inspector tabcontrol
end;

procedure TMainForm.ColorComboboxAddColors(Sender: TObject);
begin
  with TsCellCombobox(Sender) do begin
    // These are the Excel-8 palette colors, a bit rearranged and without the
    // duplicates.
    AddColor($000000, 'black');
    AddColor($333333, 'gray 80%');
    AddColor($808080, 'gray 50%');
    AddColor($969696, 'gray 40%');
    AddColor($C0C0C0, 'silver');
    AddColor($FFFFFF, 'white');
    AddColor($FF0000, 'red');
    AddColor($00FF00, 'green');
    AddColor($0000FF, 'blue');
    AddColor($FFFF00, 'yellow');
    AddColor($FF00FF, 'magenta');
    AddColor($00FFFF, 'cyan');

    AddColor($800000, 'dark red');
    AddColor($008000, 'dark green');
    AddColor($000080, 'dark blue');
    AddColor($808000, 'olive');
    AddColor($800080, 'purple');
    AddColor($008080, 'teal');
    AddColor($9999FF, 'periwinkle');
    AddColor($993366, 'plum');
    AddColor($FFFFCC, 'ivory');
    AddColor($CCFFFF, 'light turquoise');
    AddColor($660066, 'dark purple');
    AddColor($FF8080, 'coral');
    AddColor($0066CC, 'ocean blue');
    AddColor($CCCCFF, 'ice blue');

    AddColor($00CCFF, 'sky blue');
    AddColor($CCFFCC, 'light green');
    AddColor($FFFF99, 'light yellow');
    AddColor($99CCFF, 'pale blue');
    AddColor($FF99CC, 'rose');
    AddColor($CC99FF, 'lavander');
    AddColor($FFCC99, 'tan');

    AddColor($3366FF, 'light blue');
    AddColor($33CCCC, 'aqua');
    AddColor($99CC00, 'lime');
    AddColor($FFCC00, 'gold');
    AddColor($FF9900, 'light orange');
    AddColor($FF6600, 'orange');
    AddColor($666699, 'blue gray');
    AddColor($003366, 'dark teal');
    AddColor($339966, 'sea green');
    AddColor($003300, 'very dark green');
    AddColor($333300, 'olive green');
    AddColor($993300, 'brown');
    AddColor($333399, 'indigo');
  end;
end;

procedure TMainForm.EditCut1Execute(Sender: TObject);
begin
  //
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
  priorityFormats: Array[0..8] of TsSpreadFormatID;
begin
  priorityFormats[0] := ord(sfOOXML);
  priorityFormats[1] := ord(sfExcel8);
  priorityFormats[2] := ord(sfExcel5);
  priorityFormats[3] := ord(sfExcel2);
  priorityFormats[4] := ord(sfExcelXML);
  priorityFormats[5] := ord(sfOpenDocument);
  priorityFormats[6] := ord(sfCSV);
  priorityFormats[7] := sfidSYLK;
  priorityFormats[8] := ord(sfHTML);

  AcFileOpen.Dialog.Filter := GetFileFormatFilter('|', ';', faRead, priorityFormats, true, true);
  FOpenFormats := GetSpreadFormats(faRead, priorityFormats);

  AcFileSaveAs.Dialog.Filter := GetFileFormatFilter('|', ';', faWrite, priorityFormats);
  FSaveFormats := GetSpreadFormats(faWrite, priorityFormats);

 {$IFDEF WINDOWS}
  if Win32MajorVersion >= 10 then begin
    WorksheetGrid.TitleStyle := tsLazarus;
    Inspector.TitleStyle := tsLazarus;
  end;
 {$ENDIF}
end;

{ Event handler for hyperlinks: it only has to provide the hyperlink data
  which are applied to the active cell by the TsCellHyperlinkAction.
  Is called by the "new hyperlink" and "edit hyperlink" actions.
  Here we open the HyperlinkForm which is similar to the one used by
  Open/LibreOffice.

  Caption .... Caption of the form in which the hyperlink can be specified
  Hyperlink .. Data record (target, tooltip) for/from the the hyperlink form. }
procedure TMainForm.HyperlinkHandler(Sender: TObject; ACaption: String;
  var AHyperlink: TsHyperlink);
begin
  if HyperlinkForm = nil then
    HyperlinkForm := THyperlinkForm.Create(self);
  HyperlinkForm.Caption := ACaption;
  HyperlinkForm.SetHyperlink(WorkbookSource.Worksheet, AHyperlink);
  if HyperlinkForm.ShowModal = mrOK then
    HyperlinkForm.GetHyperlink(AHyperlink);
end;

{ Event handler to synchronize the mode of the spreadsheet inspector with the
  selected tab of the TabControl }
procedure TMainForm.InspectorTabControlChange(Sender: TObject);
begin
  Inspector.Mode := TsInspectorMode(InspectorTabControl.TabIndex);
end;

procedure TMainForm.ReadFromIni;
var
  ini: TCustomIniFile;
begin
  ini := CreateIni;
  try
    //...
  finally
    ini.Free;
  end;
end;

procedure TMainForm.SearchClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  Unused(CloseAction);
  DefaultSearchParams := TSearchForm(Sender).SearchParams;
  DefaultReplaceParams := TSearchForm(Sender).ReplaceParams;
  SEARCH_DLG_POS.X := SearchForm.Left;
  SEARCH_DLG_POS.Y := SearchForm.Top;
end;

procedure TMainForm.SearchFound(Sender: TObject; AFound: Boolean;
  AWorksheet: TsWorksheet; ARow, ACol: Cardinal);
begin
  Unused(AWorksheet, ARow, ACol);

  if AFound then
  begin
    //
  end
  else
  begin
    DefaultSearchParams := TSearchForm(Sender).SearchParams;
    MessageDlg(
      Format('The search text "%s" could not be found.', [DefaultSearchParams.SearchText]),
      mtInformation,
      [mbOK], 0
    );
  end;
end;

procedure TMainForm.TSaveDialogTypeChange(Sender: TObject);
var
  ext: String;
begin
  ext := GetSpreadFormatExt(FSaveFormats[AcFileSaveAs.Dialog.FilterIndex - 1]);
  AcFileSaveAs.Dialog.FileName := ChangeFileExt(AcFileSaveAs.Dialog.FileName, ext);
end;

procedure TMainForm.UpdateCaption;
begin
  if WorkbookSource = nil then
    Caption := 'demo_ctrls'
  else
    Caption := Format('demo_ctrls - "%s" [%s]', [
      AnsiToUTF8(WorkbookSource.Filename),
      GetSpreadTechnicalName(WorkbookSource.Workbook.FileFormatID)
    ]);
end;

procedure TMainForm.WriteToIni;
var
  ini: TCustomIniFile;
begin
  ini := CreateIni;
  try
    // ...
  finally
    ini.Free;
  end;
end;

{ Event handler if an external hyperlink in a cell is activated. Usually the
  linked documents/web sites etc. are opened. }
procedure TMainForm.WorksheetGridClickHyperlink(Sender: TObject;
  const AHyperlink: TsHyperlink);
var
  u: TUri;
begin
  u := ParseURI(AHyperlink.Target);
  case Lowercase(u.Protocol) of
    'http', 'https', 'ftp', 'mailto', 'file':
      OpenUrl(AHyperlink.Target);
    else
      ShowMessage('Hyperlink ' + AHyperlink.Target + ' clicked');
  end;
end;

procedure TMainForm.WorksheetGridMouseWheel(Sender: TObject;
  Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
const
  GROWTH_FACTOR = 1.05;
begin
  if ([ssCtrl, ssShift] * Shift = [ssCtrl, ssShift]) then begin
    if WheelDelta > 0 then
      WorksheetGrid.ZoomFactor := GROWTH_FACTOR* WorksheetGrid.ZoomFactor
    else
      WorksheetGrid.ZoomFactor := WorksheetGrid.ZoomFactor / GROWTH_FACTOR;
    Handled := true;
  end;
end;

end.

