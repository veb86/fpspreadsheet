unit fpsActions;

interface

uses
  SysUtils, Classes, Controls, Graphics, ActnList, StdActns, Dialogs,
  fpspreadsheet, fpspreadsheetctrls;

type
  TsSpreadsheetAction = class(TCustomAction)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetSelection: TsCellRangeArray;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    property Selection: TsCellRangeArray read GetSelection;
    property Worksheet: TsWorksheet read GetWorksheet;
  public
    function HandlesTarget(Target: TObject): Boolean; override;
    procedure UpdateTarget(Target: TObject); override;
    property Workbook: TsWorkbook read GetWorkbook;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write FWorkbookSource;
  end;

  { --- Actions related to worksheets --- }

  TsWorksheetAction = class(TsSpreadsheetAction)
  private
  public
    function HandlesTarget(Target: TObject): Boolean; override;
    procedure UpdateTarget(Target: TObject); override;
    property Worksheet;
  published
    property Caption;
    property Enabled;
    property HelpContext;
    property HelpKeyword;
    property HelpType;
    property Hint;
    property ImageIndex;
    property OnExecute;
    property OnHint;
    property OnUpdate;
    property SecondaryShortCuts;
    property ShortCut;
    property Visible;
  end;

  TsWorksheetNameEvent = procedure (Sender: TObject; AWorksheet: TsWorksheet;
    var ASheetName: String) of object;

  { Action for adding a worksheet }
  TsWorksheetAddAction = class(TsWorksheetAction)
  private
    FNameMask: String;
    FOnGetWorksheetName: TsWorksheetNameEvent;
    procedure SetNameMask(const AValue: String);
  protected
    function GetUniqueSheetName: String;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  published
    property NameMask: String read FNameMask write SetNameMask;
    property OnGetWorksheetName: TsWorksheetNameEvent
      read FOnGetWorksheetName write FOnGetWorksheetName;
  end;

  { Action for deleting selected worksheet }
  TsWorksheetDeleteAction = class(TsWorksheetAction)
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  end;

  { Action for renaming selected worksheet }
  TsWorksheetRenameAction = class(TsWorksheetAction)
  private
    FOnGetWorksheetName: TsWorksheetNameEvent;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  published
    property OnGetWorksheetName: TsWorksheetNameEvent
      read FOnGetWorksheetName write FOnGetWorksheetName;
  end;


  { --- Actions related to cell and cell selection formatting--- }

  TsCellFormatAction = class(TsSpreadsheetAction)
  private
    //
  protected
    procedure ApplyFormatToCell(ACell: PCell); virtual;
    procedure ApplyFormatToRange(ARange: TsCellrange); virtual;
    procedure ApplyFormatToSelection; virtual;
    procedure ExtractFromCell(ACell: PCell); virtual;
  public
    procedure ExecuteTarget(Target: TObject); override;
    function HandlesTarget(Target: TObject): Boolean; override;
    procedure UpdateTarget(Target: TObject); override;
    property Selection;
    property Worksheet;
  published
    property Caption;
    property Enabled;
    property HelpContext;
    property HelpKeyword;
    property HelpType;
    property Hint;
    property ImageIndex;
    property OnExecute;
    property OnHint;
    property OnUpdate;
    property SecondaryShortCuts;
    property ShortCut;
    property Visible;
  end;


  { TsFontStyleAction }

  TsFontStyleAction = class(TsCellFormatAction)
  private
    FFontStyle: TsFontStyle;
    procedure SetFontStyle(AValue: TsFontStyle);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property FontStyle: TsFontStyle
      read FFontStyle write SetFontStyle;
  end;


  { TsHorAlignmentAction }

  TsHorAlignmentAction = class(TsCellFormatAction)
  private
    FHorAlign: TsHorAlignment;
    procedure SetHorAlign(AValue: TsHorAlignment);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property HorAlignment: TsHorAlignment
      read FHorAlign write SetHorAlign default haDefault;
  end;


  { TsVertAlignmentAction }

  TsVertAlignmentAction = class(TsCellFormatAction)
  private
    FVertAlign: TsVertAlignment;
    procedure SetVertAlign(AValue: TsVertAlignment);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property VertAlignment: TsVertAlignment
      read FVertAlign write SetVertAlign default vaDefault;
  end;


  { TsTextRotationAction }

  TsTextRotationAction = class(TsCellFormatAction)
  private
    FTextRotation: TsTextRotation;
    procedure SetTextRotation(AValue: TsTextRotation);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property TextRotation: TsTextRotation
      read FTextRotation write SetTextRotation default trHorizontal;
  end;


  { TsWordwrapAction }

  TsWordwrapAction = class(TsCellFormatAction)
  private
    function GetWordwrap: Boolean;
    procedure SetWordwrap(AValue: Boolean);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Wordwrap: boolean
      read GetWordwrap write SetWordwrap default false;
  end;


  { TsNumberFormatAction }

  TsNumberFormatAction = class(TsCellFormatAction)
  private
    FNumberFormat: TsNumberFormat;
    FNumberFormatStr: string;
    procedure SetNumberFormat(AValue: TsNumberFormat);
    procedure SetNumberFormatStr(AValue: String);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property NumberFormat: TsNumberFormat
      read FNumberFormat write SetNumberFormat default nfGeneral;
    property NumberFormatString: string
      read FNumberFormatStr write SetNumberFormatStr;
  end;


  { TsDecimalsAction }
  TsDecimalsAction = class(TsCellFormatAction)
  private
    FDecimals: Integer;
    FDelta: Integer;
    procedure SetDelta(AValue: Integer);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Caption stored false;
    property Delta: Integer
      read FDelta write SetDelta default +1;
    property Hint stored false;
  end;


  { --- Actions like from TCommonDialogAction --- }

  TsCommonDialogSpreadsheetAction = class(TsCellFormatAction)
  private
    FBeforeExecute: TNotifyEvent;
    FExecuteResult: Boolean;
    FOnAccept: TNotifyEvent;
    FOnCancel: TNotifyEvent;
  protected
    FDialog: TCommonDialog;
    procedure DoAccept; virtual;
    procedure DoBeforeExecute; virtual;
    procedure DoCancel; virtual;
    function GetDialogClass: TCommonDialogClass; virtual;
    procedure CreateDialog; virtual;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
    property ExecuteResult: Boolean read FExecuteResult;
    property BeforeExecute: TNotifyEvent read FBeforeExecute write FBeforeExecute;
    property OnAccept: TNotifyEvent read FOnAccept write FOnAccept;
    property OnCancel: TNotifyEvent read FOnCancel write FOnCancel;
  end;

  { TsFontAction }
  TsFontAction = class(TsCommonDialogSpreadsheetAction)
  private
    function GetDialog: TFontDialog;
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure DoAccept; override;
    procedure ExtractFromCell(ACell: PCell); override;
    function GetDialogClass: TCommonDialogClass; override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Caption;
    property Dialog: TFontDialog read GetDialog;
    property Enabled;
    property HelpContext;
    property HelpKeyword;
    property HelpType;
    property Hint;
    property ImageIndex;
    property ShortCut;
    property SecondaryShortCuts;
    property Visible;
//    property BeforeExecute;
//    property OnAccept;
//    property OnCancel;
    property OnHint;
  end;

  { TsBackgroundColorAction }
  TsBackgroundColorAction = class(TsCommonDialogSpreadsheetAction)
  private
    FBackgroundColor: TsColor;
    function GetDialog: TColorDialog;
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure DoAccept; override;
    procedure DoBeforeExecute; override;
    procedure ExtractFromCell(ACell: PCell); override;
    function GetDialogClass: TCommonDialogClass; override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Caption;
    property Dialog: TColorDialog read GetDialog;
    property Enabled;
    property HelpContext;
    property HelpKeyword;
    property HelpType;
    property Hint;
    property ImageIndex;
    property ShortCut;
    property SecondaryShortCuts;
    property Visible;
//    property BeforeExecute;
//    property OnAccept;
//    property OnCancel;
    property OnHint;
  end;

procedure Register;


implementation

uses
  fpsutils, fpsVisualUtils;

procedure Register;
begin
  RegisterActions('FPSpreadsheet', [
    // Worksheet-releated actions
    TsWorksheetAddAction, TsWorksheetDeleteAction, TsWorksheetRenameAction,
    // Cell or cell range formatting actions
    TsFontAction, TsFontStyleAction, TsBackgroundColorAction,
    TsHorAlignmentAction, TsVertAlignmentAction,
    TsTextRotationAction, TsWordWrapAction,
    TsNumberFormatAction, TsDecimalsAction
  ], nil);
end;


{ TsSpreadsheetAction }

function TsSpreadsheetAction.GetSelection: TsCellRangeArray;
begin
  Result := Worksheet.GetSelection;
end;

function TsSpreadsheetAction.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

function TsSpreadsheetAction.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

function TsSpreadsheetAction.HandlesTarget(Target: TObject): Boolean;
begin
  Result := (Target <> nil) and (Target = FWorkbookSource);
end;

procedure TsSpreadsheetAction.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    FWorkbookSource := nil;
end;

procedure TsSpreadsheetAction.UpdateTarget(Target: TObject);
begin
  Enabled := HandlesTarget(Target);
end;


{ TsWorksheetAction }

function TsWorksheetAction.HandlesTarget(Target: TObject): Boolean;
begin
  Result := inherited HandlesTarget(Target) and (Worksheet <> nil);
end;

procedure TsWorksheetAction.UpdateTarget(Target: TObject);
begin
  Unused(Target);
  Enabled := inherited Enabled and (Worksheet <> nil);
end;


{ TsWorksheetAddAction }

constructor TsWorksheetAddAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Add';
  Hint := 'Add empty worksheet';
  FNameMask := 'Sheet%d';
end;

{ Helper procedure which creates a default worksheetname by counting a number
  up until it provides in the NameMask a unique worksheet name. }
function TsWorksheetAddAction.GetUniqueSheetName: String;
var
  i: Integer;
begin
  Result := '';
  if Workbook = nil then
    exit;

  i := 0;
  repeat
    inc(i);
    Result := Format(FNameMask, [i]);
  until Workbook.GetWorksheetByName(Result) = nil
end;

procedure TsWorksheetAddAction.ExecuteTarget(Target: TObject);
var
  sheetName: String;
begin
  if HandlesTarget(Target) then
  begin
    // Get default name of the new worksheet
    sheetName := GetUniqueSheetName;
    // If available use own procedure to specify new worksheet name
    if Assigned(FOnGetWorksheetName) then
      FOnGetWorksheetName(self, Worksheet, sheetName);
    // Check validity of worksheet name
    if not Workbook.ValidWorksheetName(sheetName) then
    begin
      MessageDlg(Format('"5s" is not a valid worksheet name.', [sheetName]), mtError, [mbOK], 0);
      exit;
    end;
    // Add new worksheet using the worksheet name.
    Workbook.AddWorksheet(sheetName);
  end;
end;

procedure TsWorksheetAddAction.SetNameMask(const AValue: String);
begin
  if AValue = FNameMask then
    exit;

  if pos('%d', AValue) = 0 then
    raise Exception.Create('Worksheet name mask must contain a %d place-holder.');

  FNameMask := AValue;
end;


{ TsWorksheetDeleteAction }

constructor TsWorksheetDeleteAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Delete...';
  Hint := 'Delete worksheet';
end;

procedure TsWorksheetDeleteAction.ExecuteTarget(Target: TObject);
begin
  if HandlesTarget(Target) then
  begin
    // Make sure that the last worksheet is not deleted - there must always be
    // at least 1 worksheet.
    if Workbook.GetWorksheetCount = 1 then
    begin
      MessageDlg('The workbook must contain at least 1 worksheet', mtError, [mbOK], 0);
      exit;
    end;

    // Confirmation dialog
    if MessageDlg(
      Format('Do you really want to delete worksheet "%s"?', [Worksheet.Name]),
      mtConfirmation, [mbYes, mbNo], 0) <> mrYes
    then
      exit;

    // Remove the worksheet; the workbookSource takes care of selecting the
    // next worksheet after deletion.
    Workbook.RemoveWorksheet(Worksheet);
  end;
end;


{ TsWorksheetRenameAction }

constructor TsWorksheetRenameAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Rename...';
  Hint := 'Rename worksheet';
end;

procedure TsWorksheetRenameAction.ExecuteTarget(Target: TObject);
var
  s: String;
begin
  if HandlesTarget(Target) then
  begin
    s := Worksheet.Name;
    // If requested, override input box by own input
    if Assigned(FOnGetWorksheetName) then
      FOnGetWorksheetName(self, Worksheet, s)
    else
      s := InputBox('Rename worksheet', 'New worksheet name', s);
    // No change
    if s = Worksheet.Name then
      exit;
    // Check validity of new worksheet name
    if Workbook.ValidWorksheetName(s) then
      Worksheet.Name := s
    else
      MessageDlg(Format('"%s" is not a valid worksheet name.', [s]), mtError, [mbOK], 0);
  end;
end;


{ TsCellFormatAction }

{ Copies the format item for which the action is responsible to the
  specified cell. Must be overridden by descendants. }
procedure TsCellFormatAction.ApplyFormatToCell(ACell: PCell);
begin
  Unused(ACell);
end;

procedure TsCellFormatAction.ApplyFormatToRange(ARange: TsCellRange);
var
  r, c: Cardinal;
  cell: PCell;
begin
  for r := ARange.Row1 to ARange.Row2 do
    for c := ARange.Col1 to ARange.Col2 do
    begin
      cell := Worksheet.GetCell(r, c);  // Use "GetCell" here to format empty cells as well
      ApplyFormatToCell(cell);  // no check for nil required because of "GetCell"
    end;
end;

procedure TsCellFormatAction.ApplyFormatToSelection;
var
  sel: TsCellRangeArray;
  range: Integer;
begin
  sel := GetSelection;
  for range := 0 to High(sel) do
    ApplyFormatToRange(sel[range]);
end;

procedure TsCellFormatAction.ExecuteTarget(Target: TObject);
begin
  ApplyFormatToSelection;
end;

{ Extracts the format item for which the action is responsible from the
  specified cell. Must be overridden by descendants. }
procedure TsCellFormatAction.ExtractFromCell(ACell: PCell);
begin
  Unused(ACell);
end;

function TsCellFormatAction.HandlesTarget(Target: TObject): Boolean;
begin
  Result := inherited HandlesTarget(Target) and (Worksheet <> nil) and (Length(GetSelection) > 0);
end;

procedure TsCellFormatAction.UpdateTarget(Target: TObject);
var
  cell: PCell;
begin
  Unused(Target);

  Enabled := inherited Enabled and (Worksheet <> nil) and (Length(GetSelection) > 0);
  if not Enabled then
    exit;

  cell := Worksheet.FindCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  ExtractFromCell(cell);
end;


{ TsFontStyleAction }

constructor TsFontStyleAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  AutoCheck := true;
end;

procedure TsFontStyleAction.ApplyFormatToCell(ACell: PCell);
var
  fnt: TsFont;
  fs: TsFontStyles;
begin
  fnt := Workbook.GetFont(ACell^.FontIndex);
  fs := fnt.Style;
  if Checked then
    Include(fs, FFontStyle)
  else
    Exclude(fs, FFontStyle);
  Worksheet.WriteFontStyle(ACell, fs);
end;

procedure TsFontStyleAction.ExtractFromCell(ACell: PCell);
var
  fnt: TsFont;
begin
  if (ACell = nil) then
    Checked := false
  else
  if (uffBold in ACell^.UsedFormattingFields) then
    Checked := (FFontStyle = fssBold)
  else
  if (uffFont in ACell^.UsedFormattingFields) then
  begin
    fnt := Workbook.GetFont(ACell^.FontIndex);
    Checked := (FFontStyle in fnt.Style);
  end else
    Checked := false;
end;

procedure TsFontStyleAction.SetFontStyle(AValue: TsFontStyle);
begin
  FFontStyle := AValue;
  case AValue of
    fssBold: begin Caption := 'Bold'; Hint := 'Bold font'; end;
    fssItalic: begin Caption := 'Italic'; Hint := 'Italic font'; end;
    fssUnderline: begin Caption := 'Underline'; Hint := 'Underlines font'; end;
    fssStrikeout: begin Caption := 'Strikeout'; Hint := 'Strike-out font'; end;
  end;
end;


{ TsHorAlignmentAction }

constructor TsHorAlignmentAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411122312;    // Date/time when this was written
  AutoCheck := true;
end;

procedure TsHorAlignmentAction.ApplyFormatToCell(ACell: PCell);
begin
  if Checked then
    Worksheet.WriteHorAlignment(ACell, FHorAlign)
  else
    Worksheet.WriteHorAlignment(ACell, haDefault);
end;

procedure TsHorAlignmentAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffHorAlign in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := ACell^.HorAlignment = FHorAlign;
end;

procedure TsHorAlignmentAction.SetHorAlign(AValue: TsHorAlignment);
begin
  FHorAlign := AValue;
  case FHorAlign of
    haLeft   : begin Caption := 'Left'; Hint := 'Left-aligned text'; end;
    haCenter : begin Caption := 'Center'; Hint := 'Centered text'; end;
    haRight  : begin Caption := 'Right'; Hint := 'Right-aligned text'; end;
    haDefault: begin Caption := 'Default'; Hint := 'Default horizontal text alignment'; end;
  end;
end;


{ TsVertAlignmentAction }

constructor TsVertAlignmentAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411122322;    // Date/time when this was written
  AutoCheck := true;
end;

procedure TsVertAlignmentAction.ApplyFormatToCell(ACell: PCell);
begin
  if Checked then
    Worksheet.WriteVertAlignment(ACell, FVertAlign)
  else
    Worksheet.WriteVertAlignment(ACell, vaDefault);
end;

procedure TsVertAlignmentAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffVertAlign in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := ACell^.VertAlignment = FVertAlign;
end;

procedure TsVertAlignmentAction.SetVertAlign(AValue: TsVertAlignment);
begin
  FVertAlign := AValue;
  case FVertAlign of
    vaTop    : begin Caption := 'Top'; Hint := 'Top-aligned text'; end;
    vaCenter : begin Caption := 'Center'; Hint := 'Vertically centered text'; end;
    vaBottom : begin Caption := 'Bottom'; Hint := 'Bottom-aligned text'; end;
    vaDefault: begin Caption := 'Default'; Hint := 'Default vertical text alignment'; end;
  end;
end;


{ TsTextRotationAction }

constructor TsTextRotationAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411141108;    // Date/time when this was written
  AutoCheck := true;
end;

procedure TsTextRotationAction.ApplyFormatToCell(ACell: PCell);
begin
  if Checked then
    Worksheet.WriteTextRotation(ACell, FTextRotation)
  else
    Worksheet.WriteTextRotation(ACell, trHorizontal);
end;

procedure TsTextRotationAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffTextRotation in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := ACell^.TextRotation = FTextRotation;
end;

procedure TsTextRotationAction.SetTextRotation(AValue: TsTextRotation);
begin
  FTextRotation := AValue;
  case FTextRotation of
    trHorizontal:
      begin Caption := 'Horizontal'; Hint := 'Horizontal text'; end;
    rt90DegreeClockwiseRotation:
      begin Caption := '90° clockwise'; Hint := '90° clockwise rotated text'; end;
    rt90DegreeCounterClockwiseRotation:
      begin Caption := '90° counter-clockwise'; Hint := '90° counter-clockwise rotated text'; end;
    rtStacked:
      begin Caption := 'Stacked'; Hint := 'Vertically stacked horizontal letters'; end;
  end;
end;


{ TsWordwrapAction }

constructor TsWordwrapAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  AutoCheck := true;
  Caption := 'Word-wrap';
  Hint := 'Word-wrapped text';
end;

procedure TsWordwrapAction.ApplyFormatToCell(ACell: PCell);
begin
  Worksheet.WriteWordwrap(ACell, Checked);
end;

procedure TsWordwrapAction.ExtractFromCell(ACell: PCell);
begin
  Checked := (ACell <> nil) and (uffWordwrap in ACell^.UsedFormattingFields);
end;

function TsWordwrapAction.GetWordwrap: Boolean;
begin
  Result := Checked;
end;

procedure TsWordwrapAction.SetWordwrap(AValue: Boolean);
begin
  Checked := AValue;
end;


{ TsNumberFormatAction }

constructor TsNumberFormatAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  GroupIndex := 1411141258;    // Date/time when this was written
  AutoCheck := true;
  Caption := 'Number format';
  Hint := 'Number format';
end;

procedure TsNumberFormatAction.ApplyFormatToCell(ACell: PCell);
var
  nf: TsNumberFormat;
  nfstr: String;
begin
  if Checked then
  begin
    nf := FNumberFormat;
    nfstr := FNumberFormatStr;
  end else
  begin
    nf := nfGeneral;
    nfstr := '';
  end;
  if IsDateTimeFormat(nf) then
    Worksheet.WriteDateTimeFormat(ACell, nf, nfstr)
  else
    Worksheet.WriteNumberFormat(ACell, nf, nfstr);
end;

procedure TsNumberFormatAction.ExtractFromCell(ACell: PCell);
begin
  if (ACell = nil) or not (uffNumberFormat in ACell^.UsedFormattingFields) then
    Checked := false
  else
    Checked := (ACell^.NumberFormat = FNumberFormat)
      and (ACell^.NumberFormatStr = FNumberFormatStr);
end;

procedure TsNumberFormatAction.SetNumberFormat(AValue: TsNumberFormat);
begin
  FNumberFormat := AValue;
  case FNumberFormat of
    nfGeneral:
      begin Caption := 'General'; Hint := 'General format'; end;
    nfFixed:
      begin Caption := 'Fixed'; Hint := 'Fixed decimals format'; end;
    nfFixedTh:
      begin Caption := 'Fixed w/thousand separator'; Hint := 'Fixed decimal count with thousand separator'; end;
    nfExp:
      begin Caption := 'Exponential'; Hint := 'Exponential format'; end;
    nfPercentage:
      begin Caption := 'Percent'; Hint := 'Percent format'; end;
    nfCurrency:
      begin Caption := 'Currency'; Hint := 'Currency format'; end;
    nfCurrencyRed:
      begin Caption := 'Currency (red)'; Hint := 'Currency format (negative values in red)'; end;
    nfShortDateTime:
      begin Caption := 'Date/time'; Hint := 'Date and time'; end;
    nfShortDate:
      begin Caption := 'Short date'; Hint := 'Short date format'; end;
    nfLongDate:
      begin Caption := 'Long date'; Hint := 'Long date format'; end;
    nfShortTime:
      begin Caption := 'Short time'; Hint := 'Short time format'; end;
    nfLongTime:
      begin Caption := 'Long time'; Hint := 'Long time foramt'; end;
    nfShortTimeAM:
      begin Caption := 'Short time AM/PM'; Hint := 'Short 12-hour time format'; end;
    nfLongTimeAM:
      begin Caption := 'Long time AM/PM'; Hint := 'Long 12-hour time format'; end;
    nfTimeInterval:
      begin Caption := 'Time interval'; Hint := 'Time interval format'; end;
    nfCustom:
      begin Caption := 'Custom'; Hint := 'User-defined custom format'; end;
  end;
end;

procedure TsNumberFormatAction.SetNumberFormatStr(AValue: String);
begin
  FNumberFormatStr := AValue;
end;


{ TsDecimalsAction }

constructor TsDecimalsAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Decimals';
  Delta := +1;
end;

procedure TsDecimalsAction.ApplyFormatToCell(ACell: PCell);
var
  decs: Integer;
begin
  if IsDateTimeFormat(ACell^.NumberFormat) then
    exit;

  if (ACell^.ContentType in [cctEmpty, cctNumber]) and (
     (not (uffNumberFormat in ACell^.UsedFormattingFields)) or
     (ACell^.NumberFormat = nfGeneral)
  ) then
    decs := Worksheet.GetDisplayedDecimals(ACell)
  else
    decs := FDecimals;
  inc(decs, FDelta);
  if decs < 0 then decs := 0;
  Worksheet.WriteDecimals(ACell, decs);
end;

procedure TsDecimalsAction.ExtractFromCell(ACell: PCell);
var
  csym: String;
  decs: Byte;
begin
  if ACell = nil then begin
    FDecimals := 2;
    exit;
  end;

  if (ACell^.ContentType in [cctEmpty, cctNumber]) and (
     (not (uffNumberFormat in ACell^.UsedFormattingFields)) or
     (ACell^.NumberFormat = nfGeneral)
  ) then
    decs := Worksheet.GetDisplayedDecimals(ACell)
  else
    Worksheet.GetNumberFormatAttributes(ACell, decs, csym);
  FDecimals := decs;
end;

procedure TsDecimalsAction.SetDelta(AValue: Integer);
begin
  FDelta := AValue;
  if FDelta > 0 then
    Hint := 'More decimal places'
  else
    Hint := 'Less decimal places';
end;


{ TsCommonDialogSpreadsheetAction }

constructor TsCommonDialogSpreadsheetAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  CreateDialog;

  DisableIfNoHandler := False;
  Enabled := True;
end;

procedure TsCommonDialogSpreadsheetAction.CreateDialog;
var
  DlgClass: TCommonDialogClass;
begin
  DlgClass := GetDialogClass;
  if Assigned(DlgClass) then
  begin
    FDialog := DlgClass.Create(Self);
    FDialog.Name := DlgClass.ClassName;
    FDialog.SetSubComponent(True);
  end;
end;

procedure TsCommonDialogSpreadsheetAction.DoAccept;
begin
  if Assigned(FOnAccept) then
    FOnAccept(Self);
end;

procedure TsCommonDialogSpreadsheetAction.DoBeforeExecute;
begin
  if Assigned(FBeforeExecute) then
    FBeforeExecute(Self);
end;

procedure TsCommonDialogSpreadsheetAction.DoCancel;
begin
  if Assigned(FOnCancel) then
    FOnCancel(Self);
end;

function TsCommonDialogSpreadsheetAction.GetDialogClass: TCommonDialogClass;
begin
  result := nil;
end;

procedure TsCommonDialogSpreadsheetAction.ExecuteTarget(Target: TObject);
begin
  DoBeforeExecute;
  FExecuteResult := FDialog.Execute;
  if FExecuteResult then
    DoAccept
  else
    DoCancel;
end;


{ TsFontAction }

constructor TsFontAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Font';
  Hint := 'Select cell font';
end;

procedure TsFontAction.ApplyFormatToCell(ACell: PCell);
var
  sfnt: TsFont;
begin
  sfnt := TsFont.Create;
  Convert_Font_to_sFont(Workbook, GetDialog.Font, sfnt);
  Worksheet.WriteFont(ACell, Workbook.AddFont(sfnt));
end;

procedure TsFontAction.DoAccept;
begin
  ApplyFormatToSelection;
end;

procedure TsFontAction.ExtractFromCell(ACell: PCell);
var
  sfnt: TsFont;
  fnt: TFont;
begin
  fnt := TFont.Create;
  try
    if (ACell = nil) then
      sfnt := Workbook.GetDefaultFont
    else
    if uffBold in ACell^.UsedFormattingFields then
      sfnt := Workbook.GetFont(1)
    else
    if uffFont in ACell^.UsedFormattingFields then
      sfnt := Workbook.GetFont(ACell^.FontIndex)
    else
      sfnt := Workbook.GetDefaultFont;
    Convert_sFont_to_Font(Workbook, sfnt, fnt);
    GetDialog.Font.Assign(fnt);
  finally
    fnt.Free;
  end;
end;

function TsFontAction.GetDialog: TFontDialog;
begin
  Result := TFontDialog(FDialog);
end;

function TsFontAction.GetDialogClass: TCommonDialogClass;
begin
  Result := TFontDialog;
end;


{ TsBackgroundColorAction }

constructor TsBackgroundColorAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := 'Backgroundcolor';
  Hint := 'Modify background color';
end;

procedure TsBackgroundColorAction.ApplyFormatToCell(ACell: PCell);
begin
  Worksheet.WritebackgroundColor(ACell, FBackgroundColor);
end;

procedure TsBackgroundColorAction.DoAccept;
begin
  FBackgroundColor := Workbook.AddColorToPalette(TsColorValue(Dialog.Color));
  ApplyFormatToSelection;
end;

procedure TsBackgroundColorAction.DoBeforeExecute;
var
  cell: PCell;
begin
  cell := Worksheet.FindCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  if (cell = nil) or not (uffBackgroundColor in cell^.UsedFormattingFields) then
    FBackgroundColor := scNotDefined
  else
    FBackgroundColor := cell^.BackgroundColor;
  Dialog.Color := Workbook.GetPaletteColor(FBackgroundColor);
end;

procedure TsBackgroundColorAction.ExtractFromCell(ACell: PCell);
begin
  //
end;

function TsBackgroundColorAction.GetDialog: TColorDialog;
begin
  Result := TColorDialog(FDialog);
end;

function TsBackgroundColorAction.GetDialogClass: TCommonDialogClass;
begin
  Result := TColorDialog;
end;


end.
