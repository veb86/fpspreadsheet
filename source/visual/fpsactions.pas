{ fpActions }

{@@ ----------------------------------------------------------------------------
  A collection of standard actions to simplify creation of menu and toolbar
  for spreadsheet applications.

AUTHORS: Werner Pamler

LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
         distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpsActions;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Classes, Controls, Graphics, ActnList, StdActns, Dialogs,
  fpstypes, fpspreadsheet, fpspreadsheetctrls;

type
  TsSpreadsheetAction = class(TCustomAction)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetActiveCell: PCell;
    function GetSelection: TsCellRangeArray;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
  protected
    procedure ApplyFormatToCell(ACell: PCell); virtual;
    procedure ApplyFormatToRange(ARange: TsCellrange); virtual;
    procedure ApplyFormatToSelection; virtual;
    procedure ExtractFromCell(ACell: PCell); virtual;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    property ActiveCell: PCell read GetActiveCell;
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

  { Action foz zooming the selected worksheet }
  TsWorksheetZoomAction = class(TsWorksheetAction)
  private
    FZoom: Integer;
    procedure SetZoom(AValue: Integer);
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
    procedure UpdateTarget({%H-}Target: TObject); override;
  published
    property Zoom: Integer read FZoom write SetZoom default 100;
  end;


  { --- Actions related to cell and cell selection formatting--- }

  TsCopyItem = (ciFormat, ciValue, ciFormula, ciAll);
  TsCopyMode = (cmBrush, cmCopy, cmCut, cmPaste);

  TsCopyAction = class(TsSpreadsheetAction)
  private
    FCopyItem: TsCopyItem;
    FCopyMode: TsCopyMode;
  public
    procedure ExecuteTarget(Target: TObject); override;
    function HandlesTarget(Target: TObject): Boolean; override;
    procedure UpdateTarget(Target: TObject); override;
  published
    property Caption;
    property CopyItem: TsCopyItem read FCopyItem write FCopyItem default ciFormat;
    property CopyMode: TsCopyMode read FCopyMode write FCopyMode default cmBrush;
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

  TsCellAction = class(TsSpreadsheetAction)
  public
    function HandlesTarget(Target: TObject): Boolean; override;
    property ActiveCell;
    property Selection;
    property Worksheet;
  end;

  TsAutoFormatAction = class(TsCellAction)
  protected
    procedure UpdateTargetFromActiveCell;
  public
    procedure ExecuteTarget(Target: TObject); override;
    procedure UpdateTarget(Target: TObject); override;
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

  { TsClearFormatAction }
  TsClearFormatAction = class(TsAutoFormatAction)
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  end;

  { TsFontStyleAction }
  TsFontStyleAction = class(TsAutoFormatAction)
  private
    FFontStyle: TsFontStyle;
    procedure SetFontStyle(AValue: TsFontStyle);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property FontStyle: TsFontStyle read FFontStyle write SetFontStyle;
  end;

  { TsHorAlignmentAction }
  TsHorAlignmentAction = class(TsAutoFormatAction)
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
  TsVertAlignmentAction = class(TsAutoFormatAction)
  private
    FVertAlign: TsVertAlignment;
    procedure SetVertAlign(AValue:TsVertAlignment);
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
  TsTextRotationAction = class(TsAutoFormatAction)
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
  TsWordwrapAction = class(TsAutoFormatAction)
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
  TsNumFormatStrEvent = procedure (Sender: TObject; AWorkbook: TsWorkbook;
    var ANumFormatStr: String) of object;

  TsNumberFormatAction = class(TsAutoFormatAction)
  private
    FNumberFormat: TsNumberFormat;
    FNumberFormatStr: string;
    FCustomNumberFormatStr: String;
    FOnGetNumFormatStr: TsNumFormatStrEvent;
    procedure SetNumberFormat(AValue: TsNumberFormat);
    procedure SetNumberFormatStr(AValue: String);
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
    procedure UpdateCaptionAndHint;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
  published
    property NumberFormat: TsNumberFormat
      read FNumberFormat write SetNumberFormat default nfGeneral;
    property NumberFormatString: string
      read FNumberFormatStr write SetNumberFormatStr;
    property OnGetNumberFormatString: TsNumFormatStrEvent
      read FOnGetNumFormatStr write FOnGetNumFormatStr;
  end;

  { TsDecimalsAction }
  TsDecimalsAction = class(TsAutoFormatAction)
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
    property Delta: Integer
      read FDelta write SetDelta default +1;
    property Hint stored false;
  end;

  { TsCellProtectionAction }
  TsCellProtectionAction = class(TsAutoFormatAction)
  private
    FProtection: TsCellProtection;
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
    procedure UpdateTarget(Target: TObject); override;
  published
    property AutoCheck default true;
    property Protection: TsCellProtection read FProtection write FProtection default cpLockCell;
  end;

  { TsCellBorderAction }
  TsActionBorders = class;

  TsActionBorder = class(TPersistent)
  private
    FOwner: TsActionBorders;
    FLineStyle: TsLineStyle;
    FColor: TColor;
    FVisible: Boolean;
    procedure SetLineStyle(AVAlue: TsLineStyle);
  protected
    procedure Changed;
  public
    constructor Create(AOwner: TsActionBorders);
    function GetStyle: TsCellBorderStyle;
    procedure SetStyle(const AValue: TsCellBorderStyle);
    procedure SetVisible(const AValue: Boolean);
  published
    property LineStyle: TsLineStyle read FLineStyle write SetLineStyle;
    property Color: TColor read FColor write FColor;
    property Visible: Boolean read FVisible write SetVisible;
  end;

  TsActionBorders = class(TPersistent)
  private
    FAction: TsCellAction;
    FBorders: Array[0..ord(High(TsCellBorder))+2] of TsActionBorder;
    function GetBorder(AIndex: Integer): TsActionBorder;
    procedure SetBorder(AIndex: integer; AValue: TsActionBorder);
  protected
    procedure Update;
  public
    constructor Create(Action: TsCellAction);
    destructor Destroy; override;
    procedure ExtractFromCell(AWorkbook: TsWorkbook; ACell: PCell);
  published
    property East: TsActionBorder index ord(cbEast)
      read GetBorder write SetBorder;
    property North: TsActionBorder index ord(cbNorth)
      read GetBorder write SetBorder;
    property South: TsActionBorder index ord(cbSouth)
      read GetBorder write SetBorder;
    property West: TsActionBorder index ord(cbWest)
      read GetBorder write SetBorder;
    property DiagonalUp: TsActionBorder index ord(cbDiagUp)
      read GetBorder write SetBorder;
    property DiagonalDown: TsActionBorder index ord(cbDiagDown)
      read GetBorder write SetBorder;
    property InnerHor: TsActionBorder index ord(High(TsCellBorder))+1
      read GetBorder write SetBorder;
    property InnerVert: TsActionBorder index ord(High(TsCellBorder))+2
      read GetBorder write SetBorder;
  end;

  TsCellBorderAction = class(TsCellAction)
  private
    FBorders: TsActionBorders;
  protected
    procedure ApplyFormatToRange(ARange: TsCellRange); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ExecuteTarget(Target: TObject); override;
  published
    property Borders: TsActionBorders read FBorders write FBorders;
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

  TsNoCellBordersAction = class(TsCellAction)
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
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


  { Cell comment actions }

  TsCellCommentMode = (ccmNew, ccmEdit, ccmDelete);

  TsCellCommentAction = class(TsCellAction)
  private
    FMode: TsCellCommentMode;
    procedure SetMode(AValue: TsCellCommentMode);
  protected
    function EditComment(ACaption: String; var AText: String): Boolean; virtual;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
    procedure UpdateTarget(Target: TObject); override;
  published
    property Mode: TsCellCommentMode read FMode write SetMode default ccmNew;
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


  { Cell hyperlink actions }

  TsCellHyperlinkMode = (chmNew, chmEdit, chmDelete);

  TsCellHyperlinkEvent = procedure (Sender: TObject; ACaption: String;
    var AHyperlink: TsHyperlink) of object;

  TsCellHyperlinkAction = class(TsCellAction)
  private
    FMode: TsCellHyperlinkMode;
    FOnHyperlink: TsCellHyperlinkEvent;
    procedure SetMode(AValue: TsCellHyperlinkMode);
  protected
    function EditHyperlink(ACaption: String; var AHyperlink: TsHyperlink): Boolean; virtual;
  public
    constructor Create(AOwner: TComponent); override;
    procedure ExecuteTarget(Target: TObject); override;
    procedure UpdateTarget(Target: TObject); override;
  published
    property Mode: TsCellHyperlinkMode read FMode write SetMode;
    property OnHyperlink: TsCellHyperlinkEvent read FOnHyperlink write FOnHyperlink;
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

  { TsMergeAction }
  TsMergeAction = class(TsAutoFormatAction)
  private
    function GetMerged: Boolean;
    procedure SetMerged(AValue: Boolean);
  protected
    procedure ApplyFormatToRange(ARange: TsCellRange); override;
    procedure ExtractFromCell(ACell: PCell); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Merged: Boolean read GetMerged write SetMerged default false;
  end;

  { --- Actions like those derived from TCommonDialogAction --- }

  TsCommonDialogSpreadsheetAction = class(TsSpreadsheetAction)
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

  { TsCommondDialogCellAction }
  TsCommonDialogCellAction = class(TsCommondialogSpreadsheetAction)
  protected
    procedure DoAccept; override;
    procedure DoBeforeExecute; override;
  public
    property ActiveCell;
    property Selection;
    property Workbook;
    property Worksheet;
  published
    property Caption;
    property Enabled;
    property HelpContext;
    property HelpKeyword;
    property HelpType;
    property Hint;
    property ImageIndex;
    property ShortCut;
    property SecondaryShortCuts;
    property Visible;
    property BeforeExecute;
    property OnAccept;
    property OnCancel;
    property OnHint;
  end;

  { TsFontDialogAction }
  TsFontDialogAction = class(TsCommonDialogCellAction)
  private
    function GetDialog: TFontDialog;
  protected
    procedure ApplyFormatToCell(ACell: PCell); override;
    procedure ExtractFromCell(ACell: PCell); override;
    function GetDialogClass: TCommonDialogClass; override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Dialog: TFontDialog read GetDialog;
  end;

  { TsBackgroundColorDialogAction }
  TsBackgroundColorDialogAction = class(TsCommonDialogCellAction)
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
    property Dialog: TColorDialog read GetDialog;
  end;


implementation

uses
  StdCtrls, ExtCtrls, Buttons, Forms,
  fpsStrings, fpsUtils, fpsNumFormat, fpsVisualUtils, fpSpreadsheetGrid;

  
{ TsSpreadsheetAction }

{ Copies the format item for which the action is responsible to the
  specified cell. Must be overridden by descendants. }
procedure TsSpreadsheetAction.ApplyFormatToCell(ACell: PCell);
begin
  Unused(ACell);
end;

procedure TsSpreadsheetAction.ApplyFormatToRange(ARange: TsCellRange);
var
  r, c, r1, c1, r2, c2: Cardinal;
  cell, base: PCell;
begin

  r := ARange.Row1;
  while r <= ARange.Row2 do
  begin
    c := ARange.Col1;
    while c <= ARange.Col2 do
    begin
      cell := Worksheet.GetCell(r, c);   // use "GetCell" here to format empty cells as well
      if Worksheet.IsMerged(cell) then begin
        Worksheet.FindMergedRange(cell, r1, c1, r2, c2);
        base := Worksheet.FindCell(r1, c1);
        ApplyFormatToCell(base);
        c := c2+1;
      end else
      begin
        ApplyFormatToCell(cell);
        inc(c);
      end;
    end;
    inc(r);
  end;
      {

  for r := ARange.Row1 to ARange.Row2 do
    for c := ARange.Col1 to ARange.Col2 do
    begin
      cell := Worksheet.GetCell(r, c);  // Use "GetCell" here to format empty cells as well
      if Worksheet.IsMerged(cell) then
      ApplyFormatToCell(cell);  // no check for nil required because of "GetCell"
    end;
    }
end;

procedure TsSpreadsheetAction.ApplyFormatToSelection;
var
  sel: TsCellRangeArray;
  range: Integer;
begin
  sel := GetSelection;
  for range := 0 to High(sel) do
    ApplyFormatToRange(sel[range]);
end;

{ Extracts the format item for which the action is responsible from the
  specified cell. Must be overridden by descendants. }
procedure TsSpreadsheetAction.ExtractFromCell(ACell: PCell);
begin
  Unused(ACell);
end;

function TsSpreadsheetAction.GetActiveCell: PCell;
begin
  if Worksheet <> nil then
    Result := Worksheet.FindCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol)
  else
    Result := nil;
end;

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
  Caption := rsAddCaption;
  Hint := rsAddWorksheetHint;
  FNameMask := 'Sheet%d';
end;

{ Helper procedure which creates a default worksheetname by counting a number
  up until - if inserted in the NameMask - it provides a unique worksheet name. }
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
  inherited;
  Caption := rsDeleteCaption;
  Hint := rsDeleteWorksheetHint;
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
  inherited;
  Caption := rsRenameCaption;
  Hint := rsRenameWorksheetHint;
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


{ TsWorksheetZoomAction }

constructor TsWorksheetZoomAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsZoomCaption;
  Hint := rsZoomWorksheetHint;
  FZoom := 100;
end;

procedure TsWorksheetZoomAction.ExecuteTarget(Target: TObject);
begin
  if HandlesTarget(Target) then
    Worksheet.Zoomfactor := FZoom / 100;
end;

procedure TsWorksheetZoomAction.SetZoom(AValue: Integer);
begin
  if AValue = FZoom then
    exit;

  if FZoom = 0 then
    raise Exception.Create('Zoomfactor cannot be 0.');

  FZoom := AValue;
end;

procedure TsWorksheetZoomAction.UpdateTarget(Target: TObject);
begin
  Checked := (Worksheet <> nil) and (round(Worksheet.ZoomFactor*100) = FZoom);
end;


{ TsCopyAction }

procedure TsCopyAction.ExecuteTarget(Target: TObject);
const
  OPERATIONS: array[TsCopyItem] of TsCopyOperation = (
    coCopyFormat, coCopyValue, coCopyFormula, coCopyCell
  );
begin
  if (Target is TsCustomWorksheetGrid) or (Target is TsWorkbookSource) then
//  if Target is TsCustomWorksheetGrid then
    case FCopyMode of
      cmBrush:
        begin
          Checked := true;
          WorkbookSource.SetPendingOperation(OPERATIONS[FCopyItem], Worksheet.GetSelection);
        end;
      cmCopy:
        begin
          Checked := false;
          WorkbookSource.CopyCellsToClipboard;
        end;
      cmCut:
        begin
          Checked := false;
          WorkbookSource.CutCellsToClipboard;
        end;
      cmPaste:
        begin
          Checked := false;
          WorkbookSource.PasteCellsFromClipboard(OPERATIONS[FCopyItem]);
        end;
    end
  else
  if Target is TCustomEdit then
    case FCopyMode of
      cmBrush : ;
      cmCopy  : (Target as TCustomEdit).CopyToClipboard;
      cmCut   : (Target as TCustomEdit).CutToClipboard;
      cmPaste : (Target as TCustomEdit).PasteFromClipboard;
    end;
end;

function TsCopyAction.HandlesTarget(Target: TObject): Boolean;
begin
  case FCopyMode of
    cmBrush: Result := inherited HandlesTarget(Target);
    cmCopy,
    cmCut,
    cmPaste: Result := (Target <> nil) and
               ( (Target is TsCustomWorksheetGrid) or (Target is TCustomEdit) );
  end;
end;

procedure TsCopyAction.UpdateTarget(Target: TObject);
begin
  Unused(Target);
  case FCopyMode of
    cmBrush:
      if WorkbookSource.PendingOperation = coNone then Checked := false;
    cmCopy, cmCut:
      Enabled := (Worksheet <> nil) and (Worksheet.GetSelectionCount > 0);
    cmPaste:
      Enabled := SpreadsheetFormatInClipboard;
  end;
end;


{ TsCellAction }

function TsCellAction.HandlesTarget(Target: TObject): Boolean;
begin
  Result := inherited HandlesTarget(Target) and (Worksheet <> nil) and (Length(GetSelection) > 0);
end;


{ TsAutoFormatAction - action for cell formatting which is automatically
  updated according to the current selection }

procedure TsAutoFormatAction.ExecuteTarget(Target: TObject);
begin
  Unused(Target);
  ApplyFormatToSelection;
end;

procedure TsAutoFormatAction.UpdateTarget(Target: TObject);
begin
  Unused(Target);
  Enabled := inherited Enabled and (Worksheet <> nil) and (Length(GetSelection) > 0);
  if Enabled then
    UpdateTargetFromActiveCell;
end;

procedure TsAutoFormatAction.UpdateTargetFromActiveCell;
var
  cell: PCell;
begin
  cell := ActiveCell;
  if Worksheet.IsMerged(cell) then
    cell := Worksheet.FindMergeBase(cell);
  ExtractFromCell(cell);
end;


{ TsClearFormatAction }

constructor TsClearFormatAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsClearFormat;
  Hint := rsClearFormat;
end;

procedure TsClearFormatAction.ApplyFormatToCell(ACell: PCell);
begin
  if ACell <> nil then
  begin
    ACell^.FormatIndex := 0;  // Default format
    Worksheet.ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;


{ TsFontStyleAction }

constructor TsFontStyleAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsCellFontCaption;
  Hint := rsCellFontHint;
  AutoCheck := true;
end;

procedure TsFontStyleAction.ApplyFormatToCell(ACell: PCell);
var
  fnt: TsFont;
  fs: TsFontStyles;
begin
  fnt := Worksheet.ReadCellFont(ACell);
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
  fmt: PsCellFormat;
begin
  if (ACell = nil) then
  begin
    Checked := false;
    exit;
  end;

  fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
  if (uffFont in fmt^.UsedFormattingFields) then
  begin
    fnt := Workbook.GetFont(fmt^.FontIndex);
    Checked := (FFontStyle in fnt.Style);
  end else
    Checked := false;
end;

procedure TsFontStyleAction.SetFontStyle(AValue: TsFontStyle);
begin
  if AValue = FFontStyle then
    exit;
  FFontStyle := AValue;
  case AValue of
    fssBold      : Caption := rsFontStyle_Bold;
    fssItalic    : Caption := rsFontStyle_Italic;
    fssUnderline : Caption := rsFontStyle_Underlined;
    fssStrikeOut : Caption := rsFontStyle_StrikeThrough;
  end;
  Hint := Caption;
end;


{ TsHorAlignmentAction }

constructor TsHorAlignmentAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsHorAlignment_Default;
  Hint := rsHorAlignment_Default;
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
  Checked := (ACell <> nil) and (Worksheet.ReadHorAlignment(ACell) = FHorAlign);
end;

procedure TsHorAlignmentAction.SetHorAlign(AValue: TsHorAlignment);
begin
  if AValue = FHorAlign then exit;
  FHorAlign := AValue;
  case FHorAlign of
    haDefault: Caption := rsHorAlignment_Default;
    haLeft   : Caption := rsHorAlignment_Left;
    haCenter : Caption := rsHorAlignment_Center;
    haRight  : Caption := rsHorAlignment_Right;
  end;
  Hint := Caption;
end;


{ TsVertAlignmentAction }

constructor TsVertAlignmentAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := rsVertAlignment_Default;
  Hint := rsVertAlignment_Default;
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
  Checked := (ACell <> nil) and (Worksheet.ReadVertAlignment(ACell) = FVertAlign);
end;

procedure TsVertAlignmentAction.SetVertAlign(AValue: TsVertAlignment);
begin
  if AValue = FVertAlign then exit;
  FVertAlign := AValue;
  case FVertAlign of
    vaDefault: Caption := rsVertAlignment_Default;
    vaTop    : Caption := rsVertAlignment_Top;
    vaCenter : Caption := rsVertAlignment_Center;
    vaBottom : Caption := rsVertAlignment_Bottom;
  end;
  Hint := Caption;
end;


{ TsTextRotationAction }

constructor TsTextRotationAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Caption := rsTextRotation_Hor;
  Hint := rsTextRotation_Hor;
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
  Checked := (ACell <> nil) and (Worksheet.ReadTextRotation(ACell) = FTextRotation);
end;

procedure TsTextRotationAction.SetTextRotation(AValue: TsTextRotation);
begin
  if AValue <> FTextRotation then exit;
  FTextRotation := AValue;
  case FTextRotation of
    trHorizontal                       : Caption := rsTextRotation_Hor;
    rt90DegreeClockwiseRotation        : Caption := rsTextRotation_Vert_CW;
    rt90DegreeCounterClockwiseRotation : Caption := rsTextRotation_Vert_CCW;
    rtStacked                          : Caption := rsTextRotation_Vert_Stacked;
  end;
  Hint := Caption;
end;


{ TsWordwrapAction }

constructor TsWordwrapAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  AutoCheck := true;
  Caption := rsWordWrap;
  Hint := rsWordWrap;
end;

procedure TsWordwrapAction.ApplyFormatToCell(ACell: PCell);
begin
  Worksheet.WriteWordwrap(ACell, Checked);
end;

procedure TsWordwrapAction.ExtractFromCell(ACell: PCell);
begin
  Checked := (ACell <> nil) and Worksheet.ReadWordwrap(ACell);
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
  Caption := rsNumberFormatCaption_General;
  Hint := rsNumberFormatHint_General;
  GroupIndex := 1411141258;    // Date/time when this code was written
  //AutoCheck := true;
end;

procedure TsNumberFormatAction.ApplyFormatToCell(ACell: PCell);
var
  nf: TsNumberFormat;
  nfstr: String;
begin
  if Checked then
  begin
    nf := FNumberFormat;
    if FNumberFormat = nfCustom then
      nfstr := FCustomNumberFormatStr
    else
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

procedure TsNumberFormatAction.ExecuteTarget(Target: TObject);
var
  nf: TsNumberFormat;
  nfstr: String;
begin
  if (FNumberFormat = nfCustom) and Assigned(FOnGetNumFormatStr) then
  begin
    Worksheet.ReadNumFormat(ActiveCell, nf, nfstr);
    FOnGetNumFormatStr(self, Workbook, nfstr);
    FCustomNumberFormatStr := nfstr;
    Checked := true;
  end else
    Checked := not Checked;

  inherited;
end;

procedure TsNumberFormatAction.ExtractFromCell(ACell: PCell);
var
  nf: TsNumberFormat;
  nfs: String;
begin
  Worksheet.ReadNumFormat(ACell, nf, nfs);
  Checked := (ACell <> nil) and (nf = FNumberFormat) and
    ((FNumberFormatStr = '') or (nfs = FNumberFormatStr));
end;

procedure TsNumberFormatAction.SetNumberFormat(AValue: TsNumberFormat);
begin
  FNumberFormat := AValue;
  UpdateCaptionAndHint;
end;

procedure TsNumberFormatAction.SetNumberFormatStr(AValue: String);
begin
  FNumberFormatStr := AValue;
  UpdateCaptionAndHint;
end;

procedure TsNumberFormatAction.UpdateCaptionAndHint;

  procedure CaptionAndHint(ACaption, AHint: String);
  begin
    if Caption = '' then begin
      if FNumberFormatStr <> '' then
        Caption := Format('%s (%s)', [ACaption, FNumberFormatStr])
      else
        Caption := ACaption;
    end;
    if Hint = '' then
      Hint := AHint;
  end;

begin
  case FNumberFormat of
    nfGeneral:
      CaptionAndHint(rsNumberFormatCaption_General, rsNumberFormatHint_General);
    nfFixed:
      CaptionAndHint(rsNumberFormatCaption_Fixed, rsNumberFormatHint_Fixed);
    nfFixedTh:
      CaptionAndHint(rsNumberFormatCaption_FixedTh, rsNumberFormatHint_FixedTh);
    nfExp:
      CaptionAndHint(rsNumberFormatCaption_Exp, rsNumberFormatHint_Exp);
    nfPercentage:
      CaptionAndHint(rsNumberFormatCaption_Percentage, rsNumberFormatHint_Percentage);
    nfFraction:
      CaptionAndHint(rsNumberFormatCaption_Fraction, rsNumberFormatHint_Fraction);
    nfCurrency:
      CaptionAndHint(rsNumberFormatCaption_Currency, rsNumberFormatHint_Currency);
    nfCurrencyRed:
      CaptionAndHint(rsNumberFormatCaption_CurrencyRed, rsNumberFormatHint_CurrencyRed);
    nfShortDateTime:
      CaptionAndHint(rsNumberFormatCaption_ShortDateTime, rsNumberFormatHint_ShortDateTime);
    nfShortDate:
      CaptionAndHint(rsNumberFormatCaption_ShortDate, rsNumberFormatHint_ShortDate);
    nfLongDate:
      CaptionAndHint(rsNumberFormatCaption_LongDate, rsNumberFormatHint_LongDate);
    nfShortTime:
      CaptionAndHint(rsNumberFormatCaption_ShortTime, rsNumberFormatHint_ShortTime);
    nfLongTime:
      CaptionAndHint(rsNumberFormatCaption_LongTime, rsNumberFormatHint_LongTime);
    nfShortTimeAM:
      CaptionAndHint(rsNumberFormatCaption_ShortTimeAM, rsNumberFormatHint_ShortTimeAM);
    nfLongTimeAM:
      CaptionAndHint(rsNumberFormatCaption_LongTimeAM, rsNumberFormatHint_LongTimeAM);
    nfDayMonth:
      CaptionAndHint(rsNumberFormatCaption_DayMonth, rsNumberFormatHint_DayMonth);
    nfMonthYear:
      CaptionAndHint(rsNumberFormatCaption_MonthYear, rsNumberFormatHint_MonthYear);
    nfTimeInterval:
      CaptionAndHint(rsNumberFormatCaption_TimeInterval, rsNumberFormatHint_TimeInterval);
    nfText:
      CaptionAndHint(rsNumberFormatCaption_Text, rsNumberFormatHint_Text);
    nfCustom:
      CaptionAndHint(rsNumberFormatCaption_Custom, rsNumberFormatHint_Custom);
  end;
  Hint := Caption;
end;


{ TsDecimalsAction }

constructor TsDecimalsAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDelta := +1;
  Caption := rsMoreDecimals;
  Hint := rsLessDecimals;
end;

procedure TsDecimalsAction.ApplyFormatToCell(ACell: PCell);
var
  decs: Integer;
  nf: TsNumberFormat;
  nfs: String;
begin
  Worksheet.ReadNumFormat(ACell, nf, nfs);

  if IsDateTimeFormat(nf) then
    exit;

  if (ACell^.ContentType in [cctEmpty, cctNumber]) and (nf <> nfGeneral) then
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
  nf: TsNumberFormat;
  nfs: String;
begin
  if ACell = nil then begin
    FDecimals := 2;
    exit;
  end;

  Worksheet.ReadNumFormat(ACell, nf, nfs);
  if (ACell^.ContentType in [cctEmpty, cctNumber]) and (nf <> nfGeneral) then
    decs := Worksheet.GetDisplayedDecimals(ACell)
  else
    Worksheet.GetNumberFormatAttributes(ACell, decs, csym);
  FDecimals := decs;
end;

procedure TsDecimalsAction.SetDelta(AValue: Integer);
begin
  if AValue = FDelta then exit;
  FDelta := AValue;
  if FDelta > 0 then Caption := rsMoreDecimals else Caption := rsLessDecimals;
  Hint := Caption;
end;


{ TsCellProtectionAction }

constructor TsCellProtectionAction.Create(AOwner: TComponent);
begin
  inherited;
  AutoCheck := true;
end;

procedure TsCellProtectionAction.ApplyFormatToCell(ACell: PCell);
var
  p: TsCellProtections;
begin
  if (Worksheet.IsProtected and (spCells in Worksheet.Protection)) then
    exit;

  p := Worksheet.ReadCellProtection(ACell);
  if Checked then
    Include(p, FProtection)
  else
    Exclude(p, FProtection);
  Worksheet.WriteCellProtection(ACell, p);
end;

procedure TsCellProtectionAction.ExtractFromCell(ACell: PCell);
var
  p: TsCellProtections;
begin
  if ACell = nil then begin
    Checked := FProtection = cpLockCell;
    exit;
  end;
  p := Worksheet.ReadCellProtection(ACell);
  Checked := FProtection in p;
  Enabled := not (Worksheet.IsProtected and (spCells in Worksheet.Protection));
end;

procedure TsCellProtectionAction.UpdateTarget(Target: TObject);
begin
  Unused(Target);
  UpdateTargetFromActiveCell;
end;


{ TsCellBorderAction }

constructor TsActionBorder.Create(AOwner: TsActionBorders);
begin
  inherited Create;
  FOwner := AOwner;
end;

procedure TsActionBorder.Changed;
begin
  FOwner.Update;
end;

function TsActionBorder.GetStyle: TsCellBorderStyle;
begin
  Result.LineStyle := FLineStyle;
  Result.Color := FColor and $00FFFFFF;
end;

procedure TsActionBorder.SetLineStyle(AValue: TsLineStyle);
begin
  if FLineStyle = AValue then
    exit;
  FLineStyle := AValue;
  Changed;
end;

procedure TsActionBorder.SetStyle(const AValue: TsCellBorderStyle);
var
  rgb: TColor;
begin
  rgb := ColorToRGB(AValue.Color);
  if (FLineStyle = AValue.LineStyle) and (FColor = rgb) then
    exit;
  FLineStyle := AValue.LineStyle;
  FColor := rgb;
  Changed;
end;

procedure TsActionBorder.SetVisible(const AValue: Boolean);
begin
  if FVisible = AValue then exit;
  FVisible := AValue;
  Changed;
end;

{ --- }

constructor TsActionBorders.Create(Action: TsCellAction);
var
  cb: Integer;
begin
  inherited Create;
  FAction := Action;
  for cb := 0 to High(FBorders) do
    FBorders[cb] := TsActionBorder.Create(self);
end;

destructor TsActionBorders.Destroy;
var
  cb: Integer;
begin
  for cb := High(FBorders) downto 0 do FBorders[cb].Free;
  inherited Destroy;
end;

procedure TsActionBorders.ExtractFromCell(AWorkbook: TsWorkbook; ACell: PCell);
var
  cb: Integer;
  fmt: PsCellFormat;
  sheet: TsWorksheet;
begin
  if (ACell <> nil) then begin
    sheet := TsWorksheet(ACell^.Worksheet);
    if sheet.IsMerged(ACell) then
      ACell := sheet.FindMergeBase(ACell);
    fmt := AWorkbook.GetPointerToCellFormat(ACell^.FormatIndex);
  end;
  if (ACell = nil) or not (uffBorder in fmt^.UsedFormattingFields) then
    for cb := 0 to High(FBorders)-2 do   // inner styles not needed here
    begin
      FBorders[cb].SetStyle(DEFAULT_BORDERSTYLES[TsCellBorder(cb)]);
      FBorders[cb].Visible := false;
    end
  else
    for cb := 0 to High(FBorders)-2 do   // inner styles not needed here
    begin
      FBorders[cb].SetStyle(fmt^.BorderStyles[TsCellBorder(cb)]);
      FBorders[cb].Visible := TsCellBorder(cb) in fmt^.Border;
    end;
end;

function TsActionBorders.GetBorder(AIndex: Integer): TsActionBorder;
begin
  Result := FBorders[AIndex];
end;

procedure TsActionBorders.SetBorder(AIndex: Integer; AValue: TsActionBorder);
begin
  FBorders[AIndex] := AValue;
  Update;
end;

procedure TsActionBorders.Update;
type
  TVisibleSet = set of byte;

  { LS = get LineStyle as String }
  function _LS(ALineStyle: TsLineStyle): String;
  begin
    case ALineStyle of
      lsThin: Result := rsThin;
      lsMedium: Result := rsMedium;
      lsDashed: Result := rsDashed;
      lsDotted: Result := rsDotted;
      lsThick: Result := rsThick;
      lsDouble: Result := rsDouble;
      lsHair: Result := rsHair;
      lsMediumDash: Result := rsMediumDash;
      lsDashDot: Result := rsDashDot;
      lsMediumDashDot: Result := rsMediumDashDot;
      lsDashDotDot: Result := rsDashDotDot;
      lsMediumDashDotDot: Result := rsMediumDashDotDot;
      lsSlantDashDot: Result := rsSlantDashDot;
      else raise Exception.Create('[TsActionBorders] Linetype not supported.');
    end;
  end;

  { CH = set Caption and Hint }
  procedure _CH(ACaption, AHint: String; AParam: String = '');
  begin
    if AParam = '' then begin
      FAction.Caption := ACaption;
      FAction.Hint := AHint;
    end else begin
      FAction.Caption := Format(ACaption, [AParam]);
      FAction.Hint := Format(AHint, [AParam]);
    end;
  end;

  function Check(const Visible: TVisibleSet): Boolean;
  { 0 = North, 1 = West, 2 = East, 3 = South, 4 = DiagUp, 5 = DiagDown, 6 = InnerHor, 7 = InnerVert }
  var
    i: Byte;
    iVisible, borderVisible: Boolean;
  begin
    Result := false;
    for i := 0 to 7 do begin
      iVisible := i in Visible;
      borderVisible := FBorders[i].Visible;
      if iVisible and not borderVisible then exit;
      if not iVisible and borderVisible then exit;
    end;
    Result := true;
  end;

begin
  if Check([0]) then begin // 0 = North = Top
    if FBorders[0].LineStyle = lsThin then
      _CH(rsBorderTop_Menu, rsBorderTop_Hint)
    else
      _CH(rsBorderTopFmt_Menu, rsBorderTopFmt_Hint, _LS(FBorders[0].LineStyle));
  end else
  if Check([1]) then begin // 1 = West = Left
    if FBorders[1].LineStyle = lsThin then
      _CH(rsBorderLeft_Menu, rsBorderLeft_Hint)
    else
      _CH(rsBorderLeftFmt_Menu, rsBorderLeftFmt_Hint, _LS(FBorders[1].LineStyle))
  end else
  if Check([2]) then begin // 2 = East = Right
    if FBorders[2].LineStyle = lsThin then
      _CH(rsBorderRight_Menu, rsBorderRight_Hint)
    else
      _CH(rsBorderRightFmt_Menu, rsBorderRightFmt_Hint, _LS(FBorders[2].LineStyle))
  end else
  if Check([3]) then begin // 3 = South
    if FBorders[3].LineStyle = lsThin then
      _CH(rsBorderBottom_Menu, rsBorderBottom_Hint)
    else
      _CH(rsBorderBottomFmt_Menu, rsBorderBottomFmt_Hint, _LS(FBorders[3].LineStyle))
  end else
  if Check([4]) then begin // 4 = diagonal up
    if FBorders[4].LineStyle = lsThin then
      _CH(rsBorderDiagUp, rsBorderdiagUp)
    else
      _CH(rsBorderDiagUpFmt, rsBorderDiagUpFmt, _LS(FBorders[4].LineStyle))
  end else
  if Check([5]) then begin // 5 = diagonal down
    if FBorders[5].LineStyle = lsThin then
      _CH(rsBorderDiagDown, rsBorderDiagDown)
    else
      _CH(rsBorderDiagDownFmt, rsBorderDiagDownFmt, _LS(FBorders[5].LineStyle))
  end else
  if Check([6]) then begin // 6 = Inner hor
    if FBorders[6].LineStyle = lsThin then
      _CH(rsBorderInnerHor_Menu, rsBorderInnerHor_Hint)
    else
      _CH(rsBorderInnerHorFmt_Menu, rsBorderInnerHorFmt_Hint, _LS(FBorders[6].LineStyle))
  end else
  if Check([7]) then begin // 7 = InnerVert
    if FBorders[7].LineStyle = lsThin then
      _CH(rsBorderInnerVert_Menu, rsBorderInnerVert_Hint)
    else
      _CH(rsBorderInnerVertFmt_Menu, rsBorderInnerVertFmt_Hint, _LS(FBorders[7].LineStyle))
  end else
  if Check([0, 4]) then  // top & bottom
  begin
    if (FBorders[0].LineStyle = lsThin) and (FBorders[4].LineStyle <> lsThin) then
      // top line always thin, bottom line chan vary
      _CH(rsBorderBottomFmt_Menu, rsBorderBottomFmt_Hint, _LS(FBorders[4].LineStyle))
    else
      _CH(rsBorderBottom_Menu, rsBorderBottom_Hint)
  end else
  if check([0, 3, 6]) then begin // all horizontal lines
    if (FBorders[0].LineStyle = FBorders[3].LineStyle) and
       (FBorders[3].LineStyle = FBorders[6].LineStyle)
    then
      _CH(rsBorderAllHorFmt_Menu, rsBorderAllHorFmt_Hint, _LS(FBorders[3].LineStyle))
    else
      _CH(rsBorderAllHor_Menu, rsBorderAllHor_Hint);
  end else
  if Check([1, 2, 7]) then begin // all vertical lines
    if (FBorders[1].LineStyle = FBorders[2].LineStyle) and
       (FBorders[2].LineStyle = FBorders[7].LineStyle)
    then
      _CH(rsBorderAllVertFmt_Menu, rsBorderAllVertFmt_Hint, _LS(FBorders[2].LineStyle))
    else
      _CH(rsBorderAllVert_Menu, rsBorderAllVert_Hint)
  end else
  if Check([6, 7]) then begin   // all inner lines
    if (FBorders[6].LineStyle = FBorders[7].LineStyle) then
      _CH(rsBorderAllInnerFmt_Menu, rsBorderAllInnerFmt_Hint, _LS(FBorders[6].LineStyle))
    else
      _CH(rsBorderAllInner_Menu, rsBorderAllInner_Hint)
  end else
  if Check([0, 1, 2, 3]) then begin // all outer lines, no inner
    if (FBorders[0].LineStyle = FBorders[1].LineStyle) and
       (FBorders[1].LineStyle = FBorders[2].LineStyle) and
       (FBorders[2].LineStyle = FBorders[3].LineStyle)
    then
      _CH(rsBorderAllOuterFmt_Menu, rsBorderAllOuterFmt_Hint, _LS(FBorders[0].LineStyle))
    else
      _CH(rsBorderAllOuter_Menu, rsBorderAllOuter_Hint);
  end else
  if Check([0, 1, 2, 3, 6, 7]) then begin // inner and outer lines, inner thin, outer may change
    if (FBorders[0].LineStyle = FBorders[1].LineStyle) and
       (FBorders[1].LineStyle = FBorders[2].LineStyle) and
       (FBorders[2].LineStyle = FBorders[3].LineStyle) and
       (FBorders[6].LineStyle = FBorders[7].LineStyle) and
       (FBorders[6].LineStyle = lsThin)
    then
      _CH(rsBorderAllFmt_Menu, rsBorderAllFmt_Hint, _LS(FBorders[0].LineStyle))
    else
      _CH(rsBorderAll_Menu, rsBorderAll_Hint)
  end else
    _CH(rsCellBorder, rsCellBorder);
end;

{ --- }

constructor TsCellBorderAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FBorders := TsActionBorders.Create(Self);
  Caption := rsCellBorder;
  Hint := rsCellBorder;
end;

destructor TsCellBorderAction.Destroy;
begin
  FBorders.Free;
  inherited;
end;

procedure TsCellBorderAction.ApplyFormatToRange(ARange: TsCellRange);

  procedure ShowBorder(ABorder: TsCellBorder; ACell: PCell;
    ABorderStyle: TsCellBorderStyle; AEnable: boolean);
  var
    fmt: TsCellFormat;
  begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    if AEnable then
    begin
      Include(fmt.UsedFormattingFields, uffBorder);
      Include(fmt.Border, ABorder);
      fmt.BorderStyles[ABorder] := ABorderStyle;
      ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    end else
      Exclude(fmt.UsedFormattingFields, uffBorder);
    Worksheet.ChangedCell(ACell^.Row, ACell^.Col);
  end;

  procedure ShowBorders(ABorder: TsCellBorder; AStart, AEnd, AColRow: LongInt;
    AColRowIsCol: Boolean; ABorderStyle: TsCellBorderStyle; AEnable: Boolean);
  var
    i: Integer;
    r1, c1, r2, c2: Cardinal;
    cell: PCell;
  begin
    i := AStart;
    while i <= AEnd do
    begin
      if AColRowIsCol then
      // Scan along specified column, i.e. i is row index
      begin
        cell := Worksheet.GetCell(i, AColRow);
        if Worksheet.IsMerged(cell) then
        begin
          Worksheet.FindMergedRange(cell, r1, c1, r2, c2);
          if (LongInt(r1) >= AStart) and (LongInt(r2) <= AEnd) then
          begin
            cell := Worksheet.GetCell(r1, c1);
            ShowBorder(ABorder, cell, ABorderStyle, AEnable);
            while (i <= longint(r2)) do begin
              cell := GetWorksheet.GetCell(i, AColRow);
              inc(i);
            end;
            Continue;
          end;
        end;
      end
      else
      // Scan along specified row, i.e. i is column index
      begin
        cell := Worksheet.GetCell(AColRow, i);
        if Worksheet.IsMerged(cell) then
        begin
          Worksheet.FindMergedRange(cell, r1, c1, r2, c2);
          if (longInt(c1) >= AStart) and (LongInt(c2) <= AEnd) then
          begin
            cell := Worksheet.GetCell(r1, c1);
            ShowBorder(ABorder, cell, ABorderStyle, AEnable);
            while (i <= longInt(c2)) do begin
              cell := GetWorksheet.GetCell(AColRow, i);
              inc(i);
            end;
            Continue;
          end;
        end;
      end;
      ShowBorder(ABorder, cell, ABorderStyle, AEnable);
      inc(i);
    end;
  end;

var
  r, c: LongInt;
begin
  // Top edge of range
  ShowBorders(cbNorth, ARange.Col1, ARange.Col2, ARange.Row1, false,
    Borders.North.GetStyle, Borders.North.Visible);

  // Bottom edge of range
  ShowBorders(cbSouth, ARange.Col1, ARange.Col2, ARange.Row2, false,
    Borders.South.GetStyle, Borders.South.Visible);

  // Left edge of range
  ShowBorders(cbWest, ARange.Row1, ARange.Row2, ARange.Col1, true,
    Borders.West.GetStyle, Borders.West.Visible);

  // Right edge of range
  ShowBorders(cbEast, ARange.Row1, ARange.Row2, ARange.Col2, true,
    Borders.East.GetStyle, Borders.East.Visible);

  // Inner horizontal edges
  for r := ARange.Row1 to ARange.Row2-1 do
    ShowBorders(cbSouth, ARange.Col1, ARange.Col2, r, false,
      Borders.InnerHor.GetStyle, Borders.InnerHor.Visible);
  for r := ARange.Row1+1 to ARange.Row2 do
    ShowBorders(cbNorth, ARange.Col1, ARange.Col2, r, false,
      Borders.InnerHor.GetStyle, Borders.InnerHor.Visible);

  // Inner vertical edges
  for c := ARange.Col1 to ARange.Col2-1 do
    ShowBorders(cbEast, ARange.Row1, ARange.Row2, c, true,
      Borders.InnerVert.GetStyle, Borders.InnerVert.Visible);
  for c := ARange.Col1+1 to ARange.Col2 do
    ShowBorders(cbWest, ARange.Row1, ARange.Row2, c, true,
      Borders.InnerVert.GetStyle, Borders.InnerVert.Visible);

  // Diagonal up and down lines
  for c := ARange.Col1 to ARange.Col2 do
  begin
    ShowBorders(cbDiagUp, ARange.Row1, ARange.Row2, c, true,
      Borders.DiagonalUp.GetStyle, Borders.DiagonalUp.Visible);
    ShowBorders(cbDiagDown, ARange.Row1, ARange.Row2, c, true,
      Borders.DiagonalDown.GetStyle, Borders.DiagonalDown.Visible);
  end;
end;

procedure TsCellBorderAction.ExecuteTarget(Target: TObject);
begin
  Unused(Target);
  ApplyFormatToSelection;
end;

procedure TsCellBorderAction.ExtractFromCell(ACell: PCell);
var
  EmptyCell: TCell;
begin
  if (ACell = nil) or not (uffBorder in Worksheet.ReadUsedFormatting(ACell)) then
  begin
    EmptyCell.Row := 0;  // silence the compiler...
    InitCell(EmptyCell);
    FBorders.ExtractFromCell(Workbook, @EmptyCell);
  end else
    FBorders.ExtractFromCell(Workbook, ACell);
end;


{ TsNoCellBordersAction }

constructor TsNoCellBordersAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsNoCellBorders;
  Hint := rsNoCellBorders;
end;

procedure TsNoCellBordersAction.ApplyFormatToCell(ACell: PCell);
begin
  Worksheet.WriteBorders(ACell, []);
end;

procedure TsNoCellBordersAction.ExecuteTarget(Target: TObject);
begin
  Unused(Target);
  ApplyFormatToSelection;
end;


{ TsCellCommentAction }

constructor TsCellCommentAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsCommentCaption_New;
  Hint := rsCommentHint_New;
end;

function TsCellCommentAction.EditComment(ACaption: String;
  var AText: String): Boolean;
var
  F: TForm;
  memo: TMemo;
  panel: TPanel;
  btn: TBitBtn;
begin
  F := TForm.Create(nil);
  try
    F.Caption := ACaption;
    F.Width := 400;
    F.Height := 300;
    F.Position := poMainFormCenter;
    memo := TMemo.Create(F);
    memo.Parent := F;
    memo.Align := alClient;
    memo.BorderSpacing.Around := 4;
    memo.Lines.Text := AText;
    panel := TPanel.Create(F);
    panel.Parent := F;
    panel.Align := alBottom;
    panel.Height := 44;
    panel.BevelOuter := bvNone;
    panel.Caption := '';
    btn := TBitBtn.Create(F);
    btn.Parent := panel;
    btn.Kind := bkOK;
    btn.Left := panel.ClientWidth - 2*btn.Width - 2*8;
    btn.Top := 6;
    btn.Anchors := [akTop, akRight];
    btn := TBitBtn.Create(F);
    btn.Parent := panel;
    btn.Kind := bkCancel;
    btn.Left := panel.ClientWidth - btn.Width - 8;
    btn.Top := 6;
    btn.Anchors := [akTop, akRight];
    if F.ShowModal = mrOK then
    begin
      Result := true;
      AText := memo.Lines.Text;
    end else
      Result := false;
  finally
    F.Free;
  end;
end;

procedure TsCellCommentAction.ExecuteTarget(Target: TObject);
var
  txt: String;
  cellStr: String;
begin
  Unused(Target);

  if Worksheet = nil then
    exit;

  cellstr := GetCellString(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  case FMode of
    ccmNew:
      begin
        txt := '';
        if EditComment(Format('New comment for cell %s',[cellStr]), txt) then
          Worksheet.WriteComment(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol, txt);
      end;
    ccmEdit:
      begin
        txt := Worksheet.ReadComment(ActiveCell);
        if EditComment(Format('Edit comment for cell %s', [cellStr]), txt) then
          Worksheet.WriteComment(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol, txt);
      end;
    ccmDelete:
      Worksheet.WriteComment(ActiveCell, '');
  end;
end;

procedure TsCellCommentAction.SetMode(AValue: TsCellCommentMode);

  procedure CaptionAndHint(ACaption, AHint: String);
  begin
    Caption := ACaption;
    Hint := AHint;
  end;

begin
  if AValue = FMode then exit;
  FMode := AValue;
  case FMode of
    ccmNew   : CaptionAndHint(rsCommentCaption_New, rsCommentHint_New);
    ccmEdit  : CaptionAndHint(rsCommentCaption_Edit, rsCommentHint_Edit);
    ccmDelete: CaptionAndHint(rsCommentCaption_Delete, rsCommentHint_Delete);
  end;
end;

procedure TsCellCommentAction.UpdateTarget(Target: TObject);
begin
  Unused(Target);
  case FMode of
    ccmNew:
      Enabled := (Worksheet <> nil) and (Length(GetSelection) > 0) and not Worksheet.HasComment(ActiveCell);
    ccmEdit, ccmDelete:
      Enabled := (Worksheet <> nil) and (Worksheet.ReadComment(ActiveCell) <> '');
  end;
end;


{ TsCellHyperlinkAction }

constructor TsCellHyperlinkAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsHyperlinkCaption_New;
  Hint := rsHyperlinkHint_New;
end;

function TsCellHyperlinkAction.EditHyperlink(ACaption: String;
  var AHyperlink: TsHyperlink): Boolean;
begin
  if Assigned(FOnHyperlink) then
    FOnHyperlink(Self, ACaption, AHyperlink)
  else
    AHyperlink.Target := InputBox(ACaption, 'Target', AHyperlink.Target);

  Result := AHyperlink.Target <> '';
end;

procedure TsCellHyperlinkAction.ExecuteTarget(Target: TObject);
var
  txt: String;
  cellStr: String;
  hyperlink: TsHyperlink;
begin
  Unused(Target);

  if Worksheet = nil then
    exit;

  cellstr := GetCellString(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  case FMode of
    chmNew:
      begin
        txt := Format('New hyperlink for cell %s', [cellStr]);
        hyperlink.Target := '';
        hyperlink.Tooltip := '';
        if EditHyperlink(txt, hyperlink) then
          Worksheet.WriteHyperlink(
            Worksheet.ActiveCellRow, Worksheet.ActiveCellCol,
            hyperlink.Target, hyperlink.ToolTip
          );
      end;
    chmEdit:
      begin
        hyperlink := Worksheet.ReadHyperlink(ActiveCell);
        txt := Format('Edit hyperlink for cell %s', [cellStr]);
        if EditHyperlink(txt, hyperlink) then
        begin
          Worksheet.WriteHyperlink(
            Worksheet.ActiveCellRow, Worksheet.ActiveCellCol,
            hyperlink.Target, hyperlink.ToolTip
          );
        end;
      end;
    chmDelete:
      Worksheet.WriteHyperlink(ActiveCell, '');
  end;
end;

procedure TsCellHyperlinkAction.SetMode(AValue: TsCellHyperlinkMode);

  procedure CaptionAndHint(ACaption, AHint: String);
  begin
    Caption := ACaption;
    Hint := AHint;
  end;

begin
  if AValue = FMode then exit;
  FMode := AValue;
  case FMode of
    chmNew   : CaptionAndHint(rsHyperlinkCaption_New, rsHyperlinkHint_New);
    chmEdit  : CaptionAndHint(rsHyperlinkCaption_Edit, rsHyperlinkHint_Edit);
    chmDelete: CaptionAndHint(rsHyperlinkCaption_Delete, rsHyperlinkHint_Delete);
  end;
end;

procedure TsCellHyperlinkAction.UpdateTarget(Target: TObject);
begin
  Unused(Target);
  case FMode of
    chmNew:
      Enabled := (Worksheet <> nil) and (Length(GetSelection) > 0) and not Worksheet.HasHyperlink(ActiveCell);
    chmEdit, chmDelete:
      Enabled := (Worksheet <> nil) and Worksheet.HasHyperlink(ActiveCell);
  end;
end;



{ TsMergeAction }

constructor TsMergeAction.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  AutoCheck := true;
  Caption := rsMergeUnmerge;
  Hint := rsMergeUnmerge;
end;

procedure TsMergeAction.ApplyFormatToRange(ARange: TsCellRange);
begin
  if Merged then
    Worksheet.MergeCells(ARange.Row1, ARange.Col1, ARange.Row2, ARange.Col2)
  else
    Worksheet.UnmergeCells(ARange.Row1, ARange.Col1);
end;

procedure TsMergeAction.ExtractFromCell(ACell: PCell);
begin
  Checked := (ACell <> nil) and Worksheet.IsMerged(ACell);
end;

function TsMergeAction.GetMerged: Boolean;
begin
  Result := Checked;
end;

procedure TsMergeAction.SetMerged(AValue: Boolean);
begin
  Checked := AValue;
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
var
  cell: PCell;
begin
  if Assigned(FBeforeExecute) then
    FBeforeExecute(Self);
  cell := Worksheet.FindCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  ExtractFromCell(cell);
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
  Unused(Target);
  DoBeforeExecute;
  FExecuteResult := FDialog.Execute;
  if FExecuteResult then
    DoAccept
  else
    DoCancel;
end;


{ TsCommonDialogCellAction }

procedure TsCommonDialogCellAction.DoAccept;
begin
  ApplyFormatToSelection;
  inherited;
end;

procedure TsCommonDialogCellAction.DoBeforeExecute;
begin
  inherited;
  ExtractFromCell(ActiveCell);
end;


{ TsFontDialogAction }

constructor TsFontDialogAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsCellFontCaption;
  Hint := rsCellFontHint;
end;

procedure TsFontDialogAction.ApplyFormatToCell(ACell: PCell);
var
  sfnt: TsFont;
begin
  sfnt := TsFont.Create;
  Convert_Font_to_sFont(GetDialog.Font, sfnt);
  Worksheet.WriteFont(ACell, Workbook.AddFont(sfnt));
end;

procedure TsFontDialogAction.ExtractFromCell(ACell: PCell);
var
  sfnt: TsFont;
  fnt: TFont;
  fmt: PsCellFormat;
begin
  fnt := TFont.Create;
  try
    if (ACell = nil) then
      sfnt := Workbook.GetDefaultFont
    else begin
      fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
      {
      if (uffBold in fmt^.UsedFormattingFields) then
        sfnt := Workbook.GetFont(BOLD_FONTINDEX)
      else}
      if (uffFont in fmt^.UsedFormattingFields) then
        sfnt := Workbook.GetFont(fmt^.FontIndex)
      else
        sfnt := Workbook.GetDefaultFont;
    end;
    Convert_sFont_to_Font(sfnt, fnt);
    GetDialog.Font.Assign(fnt);
  finally
    fnt.Free;
  end;
end;

function TsFontDialogAction.GetDialog: TFontDialog;
begin
  Result := TFontDialog(FDialog);
end;

function TsFontDialogAction.GetDialogClass: TCommonDialogClass;
begin
  Result := TFontDialog;
end;


{ TsBackgroundColorDialogAction }

constructor TsBackgroundColorDialogAction.Create(AOwner: TComponent);
begin
  inherited;
  Caption := rsBackgroundColorCaption;
  Hint := rsBackgroundColorHint;
end;

procedure TsBackgroundColorDialogAction.ApplyFormatToCell(ACell: PCell);
begin
  Worksheet.WritebackgroundColor(ACell, FBackgroundColor);
end;

procedure TsBackgroundColorDialogAction.DoAccept;
begin
  FBackgroundColor := ColorToRgb(Dialog.Color);
  inherited;
end;

procedure TsBackgroundColorDialogAction.DoBeforeExecute;
begin
  inherited;
  Dialog.Color := FBackgroundColor and $00FFFFFF;
end;

procedure TsBackgroundColorDialogAction.ExtractFromCell(ACell: PCell);
var
  fmt: PsCellFormat;
begin
  FBackgroundColor := scNotDefined;
  if (ACell <> nil) then begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if  (uffBackground in fmt^.UsedFormattingFields) then
    begin
      if fmt^.Background.Style = fsSolidFill then
        FBackgroundColor := fmt^.Background.FgColor
      else
        FBackgroundColor := fmt^.Background.BgColor;
    end;
  end;
end;

function TsBackgroundColorDialogAction.GetDialog: TColorDialog;
begin
  Result := TColorDialog(FDialog);
end;

function TsBackgroundColorDialogAction.GetDialogClass: TCommonDialogClass;
begin
  Result := TColorDialog;
end;


end.
