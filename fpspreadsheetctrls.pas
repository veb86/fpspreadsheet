{@@ ----------------------------------------------------------------------------
  Unit fpspreadsheetctrls implements some <b>visual controls</b> which help
  to create a spreadsheet application without writing too much code.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.

  EXAMPLE
  * Add a <i>WorkbookSource</i> component to the form.
  * Add a <i>WorksheetTabControl</i>
  * Add a <i>WorksheetGrid</i> (from unit fpspreadsheetgrid)
  * Link their <i>WorkbookSource</i> properties to the added
    <i>WorkbookSource</i> component
  * Set the property <i>FileName</i> of the </i>WorkbookSource</i> to a
    spreadsheet file.

  As a result, the <i>WorksheetTabControl</i> displays tabs for each worksheet
  in the file, and the <i>WorksheetGrid</i> displays the worksheet according
  to the selected tab.
-------------------------------------------------------------------------------}
unit fpspreadsheetctrls;

{$mode objfpc}{$H+}

interface

uses
  Classes, Graphics, SysUtils, Controls, StdCtrls, ComCtrls, ValEdit, ActnList,
  LResources,
  fpstypes, fpspalette, fpspreadsheet, {%H-}fpsAllFormats;

type
  {@@ Event handler procedure for displaying a message if an error or
    warning occurs during reading of a workbook. }
  TsWorkbookSourceErrorEvent = procedure (Sender: TObject;
    const AMsg: String) of object;

  {@@ Describes during communication between WorkbookSource and visual controls
    which kind of item has changed: the workbook, the worksheet, a cell value,
    or a cell formatting, etc. }
  TsNotificationItem = (lniWorkbook,
    lniWorksheet, lniWorksheetAdd, lniWorksheetRemoving, lniWorksheetRemove,
    lniWorksheetRename,
    lniCell, lniSelection, lniAbortSelection, lniRow); //, lniPalette);
  {@@ This set accompanies the notification between WorkbookSource and visual
    controls and describes which items have changed in the spreadsheet. }
  TsNotificationItems = set of TsNotificationItem;


  { TsWorkbookSource }

  {@@ TsWorkbookSource links a workbook to the visual spreadsheet controls and
    help to display or edit the workbook without written code. }
  TsWorkbookSource = class(TComponent)
  private
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FListeners: TFPList;
    FAutoDetectFormat: Boolean;
    FFileName: TFileName;
    FFileFormatID: TsSpreadFormatID;
    FUserFileFormatID: TsSpreadFormatID;
    FPendingSelection: TsCellRangeArray;
    FPendingOperation: TsCopyOperation;
    FControlLockCount: Integer;
    FOptions: TsWorkbookOptions;
    FOnError: TsWorkbookSourceErrorEvent;

    // Getters / setters
    function GetFileFormat: TsSpreadsheetFormat;
    procedure SetFileFormat(AValue: TsSpreadsheetFormat);
    procedure SetFileFormatID(AValue: TsSpreadFormatID);
    procedure SetFileName(const AFileName: TFileName);
    procedure SetOptions(AValue: TsWorkbookOptions);

    // Local event handlers
    procedure CellChangedHandler(Sender: TObject; ARow, ACol: Cardinal);
    procedure CellFontChangedHandler(Sender: TObject; ARow, ACol: Cardinal);
    procedure CellSelectedHandler(Sender: TObject; ARow, ACol: Cardinal);
//    procedure WorkbookChangedPaletteHandler(Sender: TObject);
    procedure WorkbookOpenedHandler(Sender: TObject);
    procedure WorksheetAddedHandler(Sender: TObject; ASheet: TsWorksheet);
    procedure WorksheetChangedHandler(Sender: TObject; ASheet: TsWorksheet);
    procedure WorksheetRemovedHandler(Sender: TObject; ASheetIndex: Integer);
    procedure WorksheetRemovingHandler(Sender: TObject; AWorksheet: TsWorksheet);
    procedure WorksheetRenamedHandler(Sender: TObject; AWorksheet: TsWorksheet);
    procedure WorksheetSelectedHandler(Sender: TObject; AWorksheet: TsWorksheet);

  protected
    procedure AbortSelection;
    procedure DoShowError(const AErrorMsg: String);
    procedure InternalCreateNewWorkbook(AWorkbook: TsWorkbook = nil);
    procedure InternalLoadFromFile(AFileName: string; AAutoDetect: Boolean;
      AFormatID: TsSpreadFormatID; AWorksheetIndex: Integer = -1);
    procedure Loaded; override;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

  public
    procedure AddListener(AListener: TComponent);
    procedure RemoveListener(AListener: TComponent);
    procedure NotifyListeners(AChangedItems: TsNotificationItems; AData: Pointer = nil);

  public
    procedure CreateNewWorkbook;

    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = -1); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormatID: TsSpreadFormatID = sfidUnknown; AWorksheetIndex: Integer = -1); overload;
    {
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AWorksheetIndex: Integer = -1); overload;
      }

    procedure SaveToSpreadsheetFile(AFileName: string;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormatID: TsSpreadFormatID;
      AOverwriteExisting: Boolean = true); overload;

    procedure DisableControls;
    procedure EnableControls;

    procedure SelectCell(ASheetRow, ASheetCol: Cardinal);
    procedure SelectWorksheet(AWorkSheet: TsWorksheet);

    procedure ExecutePendingOperation;
    procedure SetPendingOperation(AOperation: TsCopyOperation;
      const ASelection: TsCellRangeArray);

    { Clipboard }
//    function CellClipboardEmpty: Boolean;
//    procedure ClearCellClipboard;
    procedure CopyCellsToClipboard;
    procedure CutCellsToClipboard;
    procedure PasteCellsFromClipboard(AItem: TsCopyOperation; ATransposed: Boolean = false);

  public
    {@@ Workbook linked to the WorkbookSource }
    property Workbook: TsWorkbook read FWorkbook;
    {@@ Currently selected worksheet of the workbook }
    property Worksheet: TsWorksheet read FWorksheet;
    {@@ Indicates that which operation is waiting to be executed at next cell select }
    property PendingOperation: TsCopyOperation read FPendingOperation;
    {@@ File format identifier of the next spreadsheet file to be loaded by
      means of the Filename property. Not used when AutoDetectFormat is TRUE.
      Unlike the published property "FileFormat" the FileFormatID also takes
      care of user-defined formats. }
    property FileFormatID: TsSpreadFormatID read FFileFormatID write SetFileFormatID;

  published
    {@@ Automatically detects the fileformat when loading the spreadsheet file
      specified by FileName }
    property AutoDetectFormat: Boolean read FAutoDetectFormat write FAutoDetectFormat;
    {@@ File format of the next spreadsheet file to be loaded by means of the
      Filename property. Not used when AutoDetectFormat is TRUE.
      Note that if FileFormat is sfUser then the format ID must be specified at
      runtime. }
    property FileFormat: TsSpreadsheetFormat read GetFileFormat write SetFileFormat;
    {@@ Name of the loaded spreadsheet file which is loaded by assigning a file name
      to this property. Format detection is determined by the properties
      AutoDetectFormat and FileFormat. Using this property loads the file at
      design-time. }
    property FileName: TFileName read FFileName write SetFileName;
    {@@ A set of options to be transferred to the workbook, for e.g. formula
      calculation etc. }
    property Options: TsWorkbookOptions read FOptions write SetOptions;
    {@@ A message box is displayey if an error occurs during loading of a
      spreadsheet. This behavior can be replaced by means of the event OnError. }
    property OnError: TsWorkbookSourceErrorEvent read FOnError write FOnError;
  end;


const
  GUID_SpreadsheetControl = '{CBCAAE52-D29E-4D0C-A7F4-1016C873448A}';

type
  { IsSpreadsheetControl }

  {@@ Interface which allows the workbook source to notify linked controls of
    changes in the associated workbook. }
  IsSpreadsheetControl = interface [GUID_SpreadsheetControl]
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
  end;

  { TsWorkbookTabControl }

  {@@ TsWorkbookTabControl is a tab control which displays the sheets of the
    workbook currently loaded by the WorkbookSource in tabs. Selecting another
    tab is communicated to other spreadsheet controls via the WorkbookSource. }
  TsWorkbookTabControl = class(TTabControl, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FLockCount: Integer;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Change; override;
    procedure GetSheetList(AList: TStrings);
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    {@@ The worksheet names of this workbook are currently displayed as tabs of the TabControl. }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Identifies the worksheet which corresponds to the selected tab }
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    {@@ Link to the WorkbookSource which provides the data. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
  end;


  { TsCellEdit }

  {@@ TsCellEdit allows to edit the content or formula of the active cell of a
    worksheet, simular to Excel's cell editor above the cell grid. }
  TsCellEdit = class(TMemo, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FShowHTMLText: Boolean;
    function GetSelectedCell: PCell;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure ShowCell(ACell: PCell); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure EditingDone; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    {@@ Pointer to the currently active cell in the workbook. This cell is
      displayed in the control and can be edited. }
    property SelectedCell: PCell read GetSelectedCell;
    {@@ Refers to the underlying workbook to which the edited cell belongs. }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Refers to the underlying worksheet to which the edited cell belongs. }
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property ShowHTMLText: Boolean read FShowHTMLText write FShowHTMLText default true;
    {@@ Link to the WorkbookSource which provides the workbook and worksheet. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
  end;


  { TsCellIndicator }

  {@@ TsCellIndicator displays the address of the currently active cell of the
    worksheet and workbook. Editing the address allows to jump to the corresponding
    cell. }
  TsCellIndicator = class(TEdit, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure EditingDone; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    {@@ Refers to the underlying worksheet to which the edited cell belongs. }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Refers to the underlying worksheet to which the edited cell belongs. }
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    {@@ Link to the WorkbookSource which provides the workbook and worksheet. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    {@@ Inherited from TEdit, overridden to center the text in the control by default }
    property Alignment default taCenter;
  end;


  { TsCellFormatItem }

  TsCellFormatItem = (cfiFontName, cfiFontSize, cfiFontColor, cfiBackgroundColor,
    cfiBorderColor);


  { TsCellCombobox }

 {@@ TsCellCombobox is a multi-purpose combobox for selection of formatting
     items of a cell }
  TsCellCombobox = class(TCustomCombobox, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FFormatItem: TsCellFormatItem;
    FColorRectOffset: Integer;
    FColorRectWidth: Integer;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetColorRectOffset(AValue: Integer);
    procedure SetColorRectWidth(AValue: Integer);
    procedure SetFormatItem(AValue: TsCellFormatItem);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure ApplyFormatToCell(ACell: PCell); virtual;
    procedure Change; override;
    procedure DrawItem(AIndex: Integer; ARect: TRect;
      AState: TOwnerDrawState); override;
    procedure ExtractFromCell(ACell: PCell); virtual;
    function GetActiveCell: PCell;
    procedure Loaded; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure Populate; virtual;
    procedure ProcessItem;
    procedure Select; override;
    property Items stored false;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    {@@ Refers to the underlying workbook }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Refers to the underlying worksheet containing the displayed cell }
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    {@@ Identifies the cell format property to be used in the combobox }
    property CellFormatItem: TsCellFormatItem read FFormatItem write SetFormatItem;
    {@@ Margin around the color box }
    property ColorRectOffset: Integer read FColorRectOffset write SetColorRectOffset default 2;
    {@@ Width of the color box shown for the color-related format items }
    property ColorRectWidth: Integer read FColorRectWidth write SetColorRectWidth default 10;
    {@@ Link to the WorkbookSource which provides the workbook and worksheet. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;

    { inherited properties }
    property Align;
    property Anchors;
    property ArrowKeysTraverseList;
    property AutoComplete;
    property AutoCompleteText;
    property AutoDropDown;
    property AutoSelect;
    property AutoSize; // Note: windows has a fixed height in some styles
    property BidiMode;
    property BorderSpacing;
    property BorderStyle;
    property CharCase;
    property Color;
    property Constraints;
    property DragCursor;
    property DragKind;
    property DragMode;
    property DropDownCount;
    property Enabled;
    property Font;
//    property ItemHeight;
    property ItemIndex;
//    property Items;
    property ItemWidth;
    property MaxLength;
    property OnChange;
    property OnChangeBounds;
    property OnClick;
    property OnCloseUp;
    property OnContextPopup;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnDrawItem;
    property OnEndDrag;
    property OnDropDown;
    property OnEditingDone;
    property OnEnter;
    property OnExit;
    property OnGetItems;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMeasureItem;
    property OnMouseDown;
    property OnMouseEnter;
    property OnMouseLeave;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseWheel;
    property OnMouseWheelDown;
    property OnMouseWheelUp;
    property OnSelect;
    property OnStartDrag;
    property OnUTF8KeyPress;
    property ParentBidiMode;
    property ParentColor;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
//    property ReadOnly;
    property ShowHint;
    property Sorted;
//    property Style;
    property TabOrder;
    property TabStop;
    property Text;
    property Visible;
  end;


  { TsSpreadsheetInspector }

  {@@ Classification of data displayed by the SpreadsheetInspector. Each item
    can be assigned to a tab of a TabControl. }
  TsInspectorMode = (imWorkbook, imWorksheet, imCellValue, imCellProperties);

  {@@ Inspector expanded nodes }
  TsInspectorExpandedNode = (ienFormatSettings, ienPageLayout);
  TsInspectorExpandedNodes = set of TsInspectorExpandedNode;

  {@@ TsSpreadsheetInspector displays all properties of a workbook, worksheet,
    cell content and cell formatting in a way similar to the Object Inspector
    of Lazarus. }
  TsSpreadsheetInspector = class(TValueListEditor, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FMode: TsInspectorMode;
    FExpanded: TsInspectorExpandedNodes;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetExpanded(AValue: TsInspectorExpandedNodes);
    procedure SetMode(AValue: TsInspectorMode);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure DblClick; override;
    procedure DoUpdate; virtual;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure UpdateCellValue(ACell: PCell; AStrings: TStrings); virtual;
    procedure UpdateCellProperties(ACell: PCell; AStrings: TStrings); virtual;
    procedure UpdateWorkbook(AWorkbook: TsWorkbook; AStrings: TStrings); virtual;
    procedure UpdateWorksheet(ASheet: TsWorksheet; AStrings: TStrings); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    {@@ Refers to the underlying workbook which is displayed by the inspector. }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Refers to the underlying worksheet which is displayed by the inspector. }
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    {@@ Refers to the underlying worksheet from which the active cell is taken }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    {@@ Classification of data displayed by the SpreadsheetInspector. Each mode
      can be assigned to a tab of a TabControl. }
    property Mode: TsInspectorMode read FMode write SetMode;
    {@@ inherited from TValueListEditor, activates column titles and automatic
      column width adjustment by default }
    property DisplayOptions default [doColumnTitles, doAutoColResize];
    {@@ Displays subproperties }
    property ExpandedNodes: TsInspectorExpandedNodes
      read FExpanded write SetExpanded default [ienFormatSettings, ienPageLayout];
    {@@ inherited from TValueListEditor. Turns of the fixed column by default}
    property FixedCols default 0;
    {@@ inherited from TStringGrid, but not published in TValueListEditor. }
    property ExtendedColSizing;
  end;

function SpreadsheetFormatInClipboard: Boolean;

var
  ComboColors: TsPalette = nil;

procedure Register;


implementation

uses
  Types, Math, StrUtils, TypInfo, LCLType, LCLIntf, LCLProc,
  Dialogs, Forms, Clipbrd,
  fpsStrings, fpsRegFileFormats, fpsUtils, fpsNumFormat, fpsHTMLUtils,
  fpsCSV;

var
  cfBiff8Format: Integer = 0;
  cfBiff5Format: Integer = 0;
  cfHTMLFormat: Integer = 0;
  cfTextHTMLFormat: Integer = 0;
  cfCSVFormat: Integer = 0;
  { not working...
  cfOpenDocumentFormat: Integer = 0;
  cfStarObjectDescriptor: Integer = 0; }

{@@ ----------------------------------------------------------------------------
  Registers the spreadsheet components in the Lazarus component palette,
  page "FPSpreadsheet".
-------------------------------------------------------------------------------}
procedure Register;
begin
  RegisterComponents('FPSpreadsheet', [
    TsWorkbookSource, TsWorkbookTabControl,
    TsCellEdit, TsCellIndicator, TsCellCombobox,
    TsSpreadsheetInspector
  ]);
end;

{@@ ----------------------------------------------------------------------------
  Returns TRUE if the clipboard contains a format good for pasting into a
  worksheet grid.
-------------------------------------------------------------------------------}
function SpreadsheetFormatInClipboard: Boolean;
begin
  Result := Clipboard.HasFormat(cfBiff8Format) or
            Clipboard.HasFormat(cfBiff5Format) or
//            Clipboard.HasFormat(cfOpenDocumentFormat) or
            Clipboard.HasFormat(cfHTMLFormat) or
            Clipboard.HasFormat(cfTextHTMLFormat) or
            Clipboard.HasFormat(cfCSVFormat) or
            Clipboard.HasFormat(CF_TEXT);
end;

{------------------------------------------------------------------------------}
{                               TsCellList                                     }
{------------------------------------------------------------------------------}

type
  TsCellList = class(TList)
  private
    FMultipleRanges: Boolean;
    function GetCell(AIndex: Integer): PCell;
    procedure SetCell(AIndex: Integer; ACell: PCell);
  public
    destructor Destroy; override;
    function Add(ACell: PCell): Integer;
    function AddCell(ACell: PCell): Integer;
    function AddEmptyCell(ARow, ACol: Cardinal): Integer;
    procedure Clear; override;
    procedure Delete(AIndex: Integer);
    function IndexOf(ACell: PCell): Integer;
    property CellByIndex[AIndex: Integer]: PCell read GetCell write SetCell;
    property MultipleRanges: Boolean read FMultipleRanges write FMultipleRanges;
  end;

destructor TsCellList.Destroy;
begin
  Clear;
  inherited;
end;

function TsCellList.Add(ACell: PCell): Integer;
begin
  Result := AddCell(ACell);
end;

{ Adds a copy of a specific cell to the list }
function TsCellList.AddCell(ACell: PCell): Integer;
var
  cell: PCell;
begin
  if ACell = nil then
    raise Exception.Create('[TsCellList.AddCell] Cell is nil, use AddEmptyCell.');
  Result := IndexOf(ACell);
  if Result = - 1 then
  begin
    New(cell);
    cell^ := ACell^;
    Result := inherited Add(cell);
  end;
end;

{ Adds a "non-existing" cell to the list. Such a cell is nil in the worksheet.
  Here it has ContentType = cctEmpty and UsedFormattingFields = [], i.e. it is
  an empty cell without formatting. }
function TsCellList.AddEmptyCell(ARow, ACol: Cardinal): Integer;
var
  cell: PCell;
begin
  New(cell);
  InitCell(ARow, ACol, cell^);
  Result := inherited Add(cell);
end;

procedure TsCellList.Clear;
var
  i: Integer;
begin
  for i := Count-1 downto 0 do
    Delete(i);
  inherited Clear;
end;

procedure TsCellList.Delete(AIndex: Integer);
var
  cell: PCell;
begin
  cell := GetCell(AIndex);
  Dispose(cell);
  inherited Delete(AIndex);
end;

function TsCellList.GetCell(AIndex: Integer): PCell;
begin
  Result := PCell(inherited Items[AIndex]);
end;

function TsCellList.IndexOf(ACell: PCell): Integer;
var
  cell: PCell;
begin
  for Result:=0 to Count-1 do
  begin
    cell := GetCell(Result);
    if (cell^.Row = ACell^.Row) and (cell^.Col = ACell^.Col) then
      exit;
  end;
  Result := -1;
end;

procedure TsCellList.SetCell(AIndex: Integer; ACell: PCell);
var
  cell: PCell;
begin
  cell := GetCell(AIndex);
  cell^ := ACell^;
end;



{------------------------------------------------------------------------------}
{                            TsWorkbookSource                                  }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the WorkbookSource class. Creates the internal list for the
  notified ("listening") components, and creates an empty workbook.

  @param  AOwner  Component which is responsibile for destroying the
                  WorkbookSource.
-------------------------------------------------------------------------------}
constructor TsWorkbookSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FListeners := TFPList.Create;
  FFileFormatID := ord(sfExcel8);
  CreateNewWorkbook;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookSource class.
  Cleans up the of listening component list and destroys the linked workbook.
-------------------------------------------------------------------------------}
destructor TsWorkbookSource.Destroy;
var
  i: Integer;
begin
  // Tell listeners that the workbook source will no longer exist
  for i:= FListeners.Count-1 downto 0 do
    RemoveListener(TComponent(FListeners[i]));
  // Destroy listener list
  FListeners.Free;
  // Destroy the instance of the workbook
  FWorkbook.Free;
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Generates a message to the grid to abort the selection process.
  Needed when copying a format (e.g.) cannot be executed due to overlapping
  ranges. Without the message, the grid would still be in selection mode.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.AbortSelection;
begin
  NotifyListeners([lniAbortSelection], nil);
end;

{@@ ----------------------------------------------------------------------------
  Adds a component to the listener list. All these components are notified of
  changes in the workbook.

  @param  AListener  Component to be added to the listener list notified of
                     changes
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.AddListener(AListener: TComponent);
begin
  if FListeners.IndexOf(AListener) = -1 then  // Avoid duplicates
    FListeners.Add(AListener);
end;

{@@ ----------------------------------------------------------------------------
  Event handler for the OnChangeCell event of TsWorksheet which is fired whenver
  cell content or formatting changes.

  @param   Sender   Pointer to the worksheet
  @param   ARow     Row index (in sheet notation) of the cell changed
  @param   ACol     Column index (in sheet notation) of the cell changed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CellChangedHandler(Sender: TObject;
  ARow, ACol: Cardinal);
begin
  if FWorksheet <> nil then
    NotifyListeners([lniCell], FWorksheet.FindCell(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Event handler for the OnChangeFont event of TsWorksheet which is fired
  whenever a cell font changes. The listener, in particular the worksheetGrid,
  must adapt the height of non-fixed rows
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CellFontChangedHandler(Sender: TObject;
  ARow, ACol: Cardinal);
begin
  if FWorksheet <> nil then
  begin
    NotifyListeners([lniCell], Worksheet.FindCell(ARow, ACol));
    NotifyListeners([lniRow], {%H-}Pointer(PtrInt(ARow)));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Event handler for the OnSelectCell event of TsWorksheet which is fired
  whenever another cell is selected in the worksheet. Notifies the listeners
  of the changed selection.

  @param  Sender   Pointer to the worksheet
  @param  ARow     Row index (in sheet notation) of the newly selected cell
  @param  ACol     Column index (in sheet notation) of the newly selected cell
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CellSelectedHandler(Sender: TObject;
  ARow, ACol: Cardinal);
begin
  Unused(ARow, ACol);
  NotifyListeners([lniSelection]);

  if FPendingOperation <> coNone then
  begin
    ExecutePendingOperation;
    FPendingOperation := coNone;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a new empty workbook and adds a single worksheet
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CreateNewWorkbook;
begin
  InternalCreateNewWorkbook;
  FWorksheet := FWorkbook.AddWorksheet(Format(rsDefaultSheetName,[1]));
  SelectWorksheet(FWorksheet);
end;

{@@ ----------------------------------------------------------------------------
  Disables notification of listening controls
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.DisableControls;
begin
  inc(FControlLockCount);
end;

{@@ ----------------------------------------------------------------------------
  An error has occured during loading of the workbook. Shows a message box by
  default. But a different behavior can be obtained by means of the OnError
  event.

  @param  AErrorMsg  Error message text created by the workbook reader and to be
                     displayed in a messagebox or by means of the OnError
                     handler.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.DoShowError(const AErrorMsg: String);
begin
  if Assigned(FOnError) then
    FOnError(self, AErrorMsg)
  else
    MessageDlg(AErrorMsg, mtError, [mbOK], 0);
end;

{@@ ----------------------------------------------------------------------------
  Enables notification of listening controls
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.EnableControls;
begin
  dec(FControlLockCount);
end;

{@@ ----------------------------------------------------------------------------
  Executes a "pending operation"
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.ExecutePendingOperation;
var
  destSelection: TsCellRangeArray;
  srcCell, destCell: PCell;    // Pointers to source and destination cells
  i, j, k: Cardinal;
  ofsRow, ofsCol: LongInt;

  function DistinctRanges(R1, R2: TsCellRange): Boolean;
  begin
    Result := (R2.Col1 > R1.Col2) or (R1.Col1 > R2.Col2) or
              (R2.Row1 > R1.Row2) or (R1.Row1 > R2.Row2);
  end;

begin
  ofsRow := Worksheet.ActiveCellRow - FPendingSelection[0].Row1;
  ofsCol := Worksheet.ActiveCellCol - FPendingSelection[0].Col1;

  // Calculate destination ranges which begin at the active cell
  SetLength(destSelection, Length(FPendingSelection));
  for i := 0 to High(FPendingSelection) do
    destSelection[i] := TsCellRange(Rect(
      LongInt(FPendingSelection[i].Row1) + ofsRow,
      LongInt(FPendingSelection[i].Col1) + ofsCol,
      LongInt(FPendingSelection[i].Row2) + ofsRow,
      LongInt(FPendingSelection[i].Col2) + ofsCol
    ));

  // Check for intersection between source and destination ranges
  for i:=0 to High(FPendingSelection) do
    for j:=0 to High(FPendingSelection) do
      if not DistinctRanges(FPendingSelection[i], destSelection[j]) then
      begin
        MessageDlg('Source and destination selections are overlapping. Operation aborted.',
          mtError, [mbOK], 0);
        AbortSelection;
        exit;
      end;

  // Execute pending operation
  for i:=0 to High(FPendingSelection) do
    for j:=0 to FPendingSelection[i].Row2-FPendingSelection[i].Row1 do
      for k:=0 to FPendingSelection[i].Col2-FPendingSelection[i].Col1 do
      begin
        srcCell := Worksheet.FindCell(FPendingSelection[i].Row1+j, FPendingSelection[i].Col1+k);
        destCell := Worksheet.GetCell(destSelection[i].Row1+j, destSelection[i].Col1+k);
        case FPendingOperation of
          coCopyCell   : Worksheet.CopyCell(srcCell, destCell);
          coCopyFormat : Worksheet.CopyFormat(srcCell, destCell);
          coCopyFormula: Worksheet.CopyFormula(srcCell, destCell);
          coCopyValue  : Worksheet.CopyValue(srcCell, destCell);
        end;
      end;
end;

{@@ ----------------------------------------------------------------------------
  Getter for property "FileFormat
-------------------------------------------------------------------------------}
function TsWorkbookSource.GetFileFormat: TsSpreadsheetFormat;
begin
  if FFileFormatID < 0 then
    Result := sfUser
  else
    Result := TsSpreadsheetFormat(FFileFormatID);
end;

{@@ ----------------------------------------------------------------------------
  Internal helper method which creates a new workbook without sheets
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.InternalCreateNewWorkbook(AWorkbook: TsWorkbook = nil);
begin
  FreeAndNil(FWorkbook);
  FWorksheet := nil;
  if AWorkbook = nil then
    FWorkbook := TsWorkbook.Create else
    FWorkbook := AWorkbook;
  FWorkbook.OnOpenWorkbook := @WorkbookOpenedHandler;
  FWorkbook.OnAddWorksheet := @WorksheetAddedHandler;
  FWorkbook.OnChangeWorksheet := @WorksheetChangedHandler;
  FWorkbook.OnRemoveWorksheet := @WorksheetRemovedHandler;
  FWorkbook.OnRemovingWorksheet := @WorksheetRemovingHandler;
  FWorkbook.OnRenameWorksheet := @WorksheetRenamedHandler;
  FWorkbook.OnSelectWorksheet := @WorksheetSelectedHandler;
//  FWorkbook.OnChangePalette := @WorkbookChangedPaletteHandler;
  // Pass options to workbook
  SetOptions(FOptions);
end;

{@@ ----------------------------------------------------------------------------
  Internal loader for the spreadsheet file. Is called with various combinations
  of arguments from several procedures.

  @param  AFilename        Name of the spreadsheet file to be loaded
  @param  AAutoDetect      Instructs the loader to automatically detect the
                           file format from the file extension or by temporarily
                           opening the file in all available formats. Note that
                           an exception is raised in the IDE when an incorrect
                           format is tested.
  @param  AFormatID        Identifier of the spreadsheet file format assumed
                           for the loader.
                           Is ignored when AAutoDetect is false.
  @param  AWorksheetIndex  Index of the worksheet to be selected after loading.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.InternalLoadFromFile(AFileName: string;
  AAutoDetect: Boolean; AFormatID: TsSpreadFormatID;
  AWorksheetIndex: Integer = -1);
{
// this version avoids loading a defective workbook into the grid,
// but tabcontrol navigation is not working here...

var
  book: TsWorkbook;
begin
  book := TsWorkbook.Create;
  try
    if AAutoDetect then
      book.ReadFromFile(AFileName)
    else
      book.ReadFromFile(AFileName, AFormatID);

    InternalCreateNewWorkbook(book);  // book --> FWorkbook
  except
    book.Free;
    InternalCreateNewWorkbook;
    // --> empty workbook, but avoid having no worksheet
    FWorksheet := FWorkbook.AddWorksheet(Format(rsDefaultSheetName, [1]));
    DoShowError(Format(rsCannotReadFile, [AFileName]));
  end;

  if AWorksheetIndex = -1 then
  begin
    if FWorkbook.ActiveWorksheet <> nil then
      AWorksheetIndex := FWorkbook.GetWorksheetIndex(FWorkbook.ActiveWorksheet) else
      AWorksheetIndex := 0;
  end;

  SelectWorksheet(FWorkbook.GetWorkSheetByIndex(AWorksheetIndex));

  // If required, display loading error message
  if FWorkbook.ErrorMsg <> '' then
    DoShowError(FWorkbook.ErrorMsg);
end;
}

begin
  // Create a new empty workbook
  InternalCreateNewWorkbook;

  DisableControls;
  try
    // Read workbook from file and get worksheet
    try
      if AAutoDetect then
        FWorkbook.ReadFromFile(AFileName)
      else
        FWorkbook.ReadFromFile(AFileName, AFormatID);
    except
      FWorkbook.AddErrorMsg(rsCannotReadFile, [AFileName]);
      // Code executed subsequently will be a pain if there is no worksheet!
      if FWorkbook.GetWorksheetCount = 0 then
        FWorkbook.AddWorksheet(Format(rsDefaultSheetName, [1]));
    end;
  finally
    EnableControls;
  end;

  if AWorksheetIndex = -1 then
  begin
    if FWorkbook.ActiveWorksheet <> nil then
      AWorksheetIndex := FWorkbook.GetWorksheetIndex(FWorkbook.ActiveWorksheet) else
      AWorksheetIndex := 0;
  end;
  SelectWorksheet(FWorkbook.GetWorkSheetByIndex(AWorksheetIndex));

  // If required, display loading error message
  if FWorkbook.ErrorMsg <> '' then
    DoShowError(FWorkbook.ErrorMsg);
end;

{@@ ----------------------------------------------------------------------------
  Inherited method which is called after reading the WorkbookSource from the lfm
  file.
  Is overridden here to open a spreadsheet file if a file name has been assigned
  to the FileName property at design-time.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.Loaded;
begin
  inherited;
  if (FFileName <> '') then
    SetFileName(FFilename)
  else
    CreateNewWorkbook;
end;

{@@ ----------------------------------------------------------------------------
  Public spreadsheet loader to be used if file format is known.

  Call this method only for built-in file formats.

  @param  AFilename        Name of the spreadsheet file to be loaded
  @param  AFormat          Spreadsheet file format assumed for the file
  @param  AWorksheetIndex  Index of the worksheet to be selected after loading.
                           (If empty then the active worksheet is loaded)
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromSpreadsheetFile(AFileName: string;
  AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = -1);
begin
  if AFormat = sfUser then
    raise Exception.Create('[TsWorkbook.ReadFromFile] Don''t call this method for user-provided file formats.');
  LoadFromSpreadsheetFile(AFileName, ord(AFormat), AWorksheetIndex);
end;

{@@ ----------------------------------------------------------------------------
  Public spreadsheet loader to be used if file format is known.

  Call this methdd for both built-in and user-provided file formats.

  @param  AFilename        Name of the spreadsheet file to be loaded
  @param  AFormatID        Identifier of the spreadsheet file format assumed
                           for the file
  @param  AWorksheetIndex  Index of the worksheet to be selected after loading.
                           (If empty then the active worksheet is loaded)
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromSpreadsheetFile(AFileName: string;
  AFormatID: TsSpreadFormatID = sfidUnknown; AWorksheetIndex: Integer = -1);
var
  autodetect: Boolean;
begin
  autodetect := (AFormatID = sfidUnknown);
  InternalLoadFromFile(AFileName, autodetect, AFormatID, AWorksheetIndex);
end;
                                          (*
{@@ ------------------------------------------------------------------------------
  Public spreadsheet loader to be used if file format is not known. The file
  format is determined from the file extension, or - if this is holds for
  several formats (such as .xls) - by assuming a format. Note that exceptions
  are raised in the IDE if in incorrect format is tested. This does not occur
  outside the IDE.

  @param  AFilename        Name of the spreadsheet file to be loaded
  @param  AWorksheetIndex  Index of the worksheet to be selected after loading.
                           (If empty then the active worksheet is loaded)
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromSpreadsheetFile(AFileName: string;
  AWorksheetIndex: Integer = -1);
const
  sfNotNeeded = sfUnknown;
  // The parameter AFormat if InternalLoadFromFile is not needed here,
  // but the compiler wants a value...
begin
  InternalLoadFromFile(AFileName, true, sfNotNeeded, AWorksheetIndex);
end;                                        *)

{@@ ----------------------------------------------------------------------------
  Notifies listeners of workbook, worksheet, cell, or selection changes.
  The changed item is identified by the parameter AChangedItems.

  @param  AChangedItems  A set of elements lniWorkbook, lniWorksheet,
                         lniCell, lniSelection which indicate which item has
                         changed.
  @param  AData          Additional information on the change. Is used only for
                         lniCell and points to the cell having a changed value
                         or formatting.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.NotifyListeners(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
var
  j: Integer;
  I: IsSpreadsheetControl;
  C: TComponent;
begin
  for j:=0 to FListeners.Count-1 do begin
    C := TComponent(FListeners[j]);
    if C.GetInterface(GUID_SpreadsheetControl, I) then
      I.ListenerNotification(AChangedItems, AData)
    else
      raise Exception.CreateFmt('[TsWorkbookSource.NotifyListeners] Class %s is not prepared to be a spreadsheet listener.',
        [C.ClassName]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes a component from the listener list. The component is no longer
  notified of changes in workbook, worksheet or cells

  @param  AListener  Listening component to be removed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.RemoveListener(AListener: TComponent);
var
  j: Integer;
  I: IsSpreadsheetControl;
  C: TComponent;
begin
  for j:=FListeners.Count-1 downto 0 do begin
    C := TComponent(FListeners[j]);
    if C = AListener then
    begin
      FListeners.Delete(j);
      if C.GetInterface(GUID_SpreadsheetControl, I) then
        I.RemoveWorkbookSource
      else
        raise Exception.CreateFmt('Class %s not prepared for listening.',[AListener.ClassName]);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the workbook of the WorkbookSource component to a spreadsheet file.

  Call this method only for built-in file formats.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved.
  @param   AFormat            Spreadsheet file format in which the file is to be
                              saved.
  @param   AOverwriteExisting If the file already exists, it is overwritten in
                              the case of AOverwriteExisting = true, or an
                              exception is raised otherwise.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SaveToSpreadsheetFile(AFileName: String;
  AFormat: TsSpreadsheetFormat; AOverwriteExisting: Boolean = true);
begin
  if AFormat = sfUser then
    raise Exception.Create('[TsWorkbook.ReadFromFile] Don''t call this method for user-provided file formats.');
  SaveToSpreadsheetFile(AFileName, ord(AFormat), AOverwriteExisting);
end;

{@@ ----------------------------------------------------------------------------
  Writes the workbook of the WorkbookSource component to a spreadsheet file.

  Call this method for both built-in and user-provided file formats.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved.
  @param   AFormatID          Identifier of the spreadsheet file format in which
                              the file is to be saved.
  @param   AOverwriteExisting If the file already exists, it is overwritten in
                              the case of AOverwriteExisting = true, or an
                              exception is raised otherwise.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SaveToSpreadsheetFile(AFileName: String;
  AFormatID: TsSpreadFormatID; AOverwriteExisting: Boolean = true);
begin
  if FWorkbook <> nil then begin
    FWorkbook.WriteToFile(AFileName, AFormatID, AOverwriteExisting);
    FFileName := AFilename;
    FFileFormatID := AFormatID;

    // If required, display loading error message
    if FWorkbook.ErrorMsg <> '' then
      DoShowError(FWorkbook.ErrorMsg);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Saves the workbook into a file with the specified file name.
  The file format is determined automatically from the extension.
  If this file name already exists the file is overwritten
  if AOverwriteExisting is true.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved
                              If the file format is not known is is written
                              as BIFF8/XLS.
  @param   AOverwriteExisting If this file already exists it is overwritten if
                              AOverwriteExisting = true, or an exception is
                              raised if AOverwriteExisting = false.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SaveToSpreadsheetFile(AFileName: String;
  AOverwriteExisting: Boolean = true);
begin
  if FWorkbook <> nil then begin
    FWorkbook.WriteToFile(AFileName, AOverwriteExisting);

    // If required, display loading error message
    if FWorkbook.ErrorMsg <> '' then
      DoShowError(FWorkbook.ErrorMsg);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Usually called by code or from the spreadsheet grid component. The
  method identifies a cell to be "selected".
  Stores its coordinates in the worksheet ("active cell") and notifies the
  listening controls
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SelectCell(ASheetRow, ASheetCol: Cardinal);
begin
  if FWorksheet <> nil then
  begin
    FWorksheet.SelectCell(ASheetRow, ASheetCol);
    NotifyListeners([lniSelection]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Selects a worksheet and notifies the controls. This method is usually called
  by code or from the worksheet tabcontrol.

  @param  AWorksheet  Instance of the newly selected worksheet.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SelectWorksheet(AWorkSheet: TsWorksheet);
begin
  FWorksheet := AWorksheet;
  if (FWorkbook <> nil) and (FWorksheet <> nil) then
    FWorkbook.SelectWorksheet(AWorksheet);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the property "FileFormat"
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SetFileFormat(AValue: TsSpreadsheetFormat);
begin
  if AValue = sfUser then
    FFileFormatID := FUserFileFormatID else
    FFileFormatID := ord(AValue);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the (public) property "FileFormatID"
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SetFileFormatID(AValue: TsSpreadFormatID);
begin
  if AValue >= ord(High(TsSpreadsheetFormat)) then   // ">= High()" because we don't want sfUser here.
    raise Exception.Create('Illegal ID for built-in format ID');
  FFileFormatID := AValue;
  if FFileFormatID < 0 then
    FUserFileFormatID := FFileFormatID;
end;

{@@ ----------------------------------------------------------------------------
  Setter for the file name property. Loads the spreadsheet file and uses the
  values of the properties AutoDetectFormat or FileFormat.
  Useful if the spreadsheet is to be loaded already at design time.
  But note that an exception can be raised if the file format cannot be
  determined from the file extension alone.

  @param  AFileName  Name of the spreadsheet file to be loaded.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SetFileName(const AFileName: TFileName);
begin
  FFileName := AFileName;

  if AFileName = '' then
  begin
    CreateNewWorkbook;
    exit;
  end;

  if FileExists(FFileName) then
  begin
    if FAutoDetectFormat then
      LoadFromSpreadsheetFile(FFileName)
    else
    if (csDesigning in ComponentState) and (FFileFormatID < 0) then
      raise Exception.Create('[TsWorkbookSource.SetFileName] User-defined file format not allowed in design mode.')
    else
      LoadFromSpreadsheetFile(FFileName, FFileFormatID);
  end else
    raise Exception.CreateFmt(rsFileNotFound, [ExpandFileName(AFileName)]);
end;

{@@ ----------------------------------------------------------------------------
  Setter for the property Options. Copies the options of the WorkbookSource
  to the workbook
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SetOptions(AValue: TsWorkbookOptions);
begin
  FOptions := AValue;
  if Workbook <> nil then
    Workbook.Options := FOptions;
end;

{@@ ----------------------------------------------------------------------------
  Defines a "pending operation" which will be executed at next cell select.
  Source of the operation is the selection passes as a parameter.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.SetPendingOperation(AOperation: TsCopyOperation;
  const ASelection: TsCellRangeArray);
var
  i: Integer;
begin
  SetLength(FPendingSelection, Length(ASelection));
  for i:=0 to High(FPendingSelection) do
    FPendingSelection[i] := ASelection[i];
  FPendingSelection := ASelection;
  FPendingOperation := AOperation;
end;

{@@ ----------------------------------------------------------------------------
  Copies the selected cells of the worksheet to an internal list ("Clipboard").
  Note that this is not the system clipboard in the current implementation.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CopyCellsToClipboard;
var
  sel: TsCellRangeArray;
  stream: TStream;
  savedCSVParams: TsCSVParams;

  procedure CopyToClipboard(AStream: TStream; AFileFormat: TsSpreadsheetFormat;
    AClipboardFormat: Integer; AParams: TsStreamParams = []);
  begin
    if AClipboardFormat = 0 then
      exit;
    FWorkbook.CopyToClipboardStream(AStream, AFileFormat, AParams);
    Clipboard.AddFormat(AClipboardFormat, AStream);
    (AStream as TMemoryStream).Clear;
  end;

begin
  sel := FWorksheet.GetSelection;
  if Length(sel) = 0 then
    exit;

  Clipboard.Open;
  try
    Clipboard.Clear;

    stream := TMemoryStream.Create;
    try
      { -- not working...

      // Write OpenDocument format
      CopyToClipboard(stream, sfOpenDocument, cfOpenDocumentFormat);

      // Write OpenDocument's "Star Object Descriptor"
      WriteStarObjectDescriptorToStream(stream);
      if cfStarObjectDescriptor <> 0 then
        Clipboard.AddFormat(cfStarObjectDescriptor, stream);
      (stream as TMemoryStream).Clear;
      }

      // Write BIFF8 format
      CopyToClipboard(stream, sfExcel8, cfBiff8Format);

      // Then write BIFF5 format
      CopyToClipboard(stream, sfExcel5, cfBiff5Format);

      // Then write Windows HTML format
      {$IFDEF MSWINDOWS}
      CopyToClipboard(stream, sfHTML, cfHtmlFormat, [spWindowsClipboardHTML]);
      {$ENDIF}

      // Write standard html format (MIME-type "text/html")
      CopyToClipboard(stream, sfHTML, cfTextHTMLFormat);

      // Then write CSV format
      savedCSVParams := CSVParams;
      CsvParams.Delimiter := ';';
      CsvParams.AutoDetectNumberFormat := false;
      CsvParams.SheetIndex := FWorkbook.GetWorksheetIndex(FWorkbook.ActiveWorksheet);
      CopyToClipboard(stream, sfCSV, cfCSVFormat);

      // Finally write TEXT format
      CsvParams.Delimiter := #9;
      CopyToClipboard(stream, sfCSV, CF_TEXT);
      CSVParams := savedCSVParams;

      // To do: XML format, ods format
    finally
      stream.Free;
    end;
  finally
    Clipboard.Close;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Copies the selected cells of the worksheet to an internal list ("Clipboard")
  and deletes them afterwards.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CutCellsToClipboard;
begin
  CopyCellsToClipboard;
  FWorksheet.DeleteSelection;
end;

{@@ ----------------------------------------------------------------------------
  Pastes the cells stored in the internal list "Clipboard" into the worksheet.
  Using their stored row/col indexes the stored cells are translated such that
  the first stored cell appears at the currently active cell in the worksheet.

  AOperation determines which "item" of the cell (all, values, formats, formula)
  is pasted.

  If ATranspose is TRUE then rows and columns are interchanged.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.PasteCellsFromClipboard(AItem: TsCopyOperation;
  ATransposed: Boolean = false);
var
  fmt: TsSpreadsheetFormat;
  stream: TStream;
  params: TsStreamParams;
begin
  stream := TMemoryStream.Create;
  try
    params := [spClipboard];
    // Check whether the clipboard content is suitable for fpspreadsheet
    {if Clipboard.GetFormat(cfOpenDocumentFormat, stream) then
      fmt := sfOpenDocument
    else}
    if Clipboard.GetFormat(cfBiff8Format, stream) then
      fmt := sfExcel8
    else if Clipboard.GetFormat(cfBiff5Format, stream) then
      fmt := sfExcel5
    else if Clipboard.GetFormat(cfHTMLFormat, stream) then begin
      fmt := sfHTML;
      params := params + [spWindowsClipboardHTML];
    end else if Clipboard.GetFormat(cfTextHTMLFormat, stream) then
      fmt := sfHTML
    else if Clipboard.GetFormat(cfCSVFormat, stream) then
      fmt := sfCSV
    else if Clipboard.GetFormat(CF_TEXT, stream) then
      fmt := sfCSV
    else begin
      // Exit if there are no spreadsheet data in clipboard
      MessageDlg('No appropriate spreadsheet data in clipboard', mtError, [mbOk], 0);
      exit;
    end;

    // Paste stream into workbook
    stream.Position := 0;
    FWorkbook.PasteFromClipboardStream(stream, fmt, AItem, params, ATransposed);

    // To do: XML format, ods format
  finally
    stream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Event handler called whenever a new workbook is opened.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorkbookOpenedHandler(Sender: TObject);
begin
  Unused(Sender);
  NotifyListeners([lniWorkbook]);
  if FWorkbook.ActiveWorksheet = nil then
    SelectWorksheet(FWorkbook.GetFirstWorksheet) else
    SelectWorksheet(FWorkbook.ActiveWorksheet);
end;

{@@ ----------------------------------------------------------------------------
  Event handler called whenever a new worksheet is added to the workbook

  @param  Sender   Pointer to the workbook to which a new worksheet has been added
  @param  ASheet   Worksheet which is added to the workbook.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetAddedHandler(Sender: TObject;
  ASheet: TsWorksheet);
begin
  Unused(Sender);
  NotifyListeners([lniWorksheetAdd]);
  SelectWorksheet(ASheet);
end;

{@@ ----------------------------------------------------------------------------
  Event handler canned whenever worksheet properties have changed.
  Currently only used for changing the workbook name.

  @param  Sender  Workbook containing the modified worksheet
  @param  ASheet  Worksheet which has been modified
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetChangedHandler(Sender: TObject;
  ASheet: TsWorksheet);
begin
  Unused(Sender, ASheet);
  NotifyListeners([lniWorkbook, lniWorksheet]);
end;

{@@ ----------------------------------------------------------------------------
  Event handler called AFTER a worksheet has been removed (deleted) from
  the workbook

  @param  Sender       Points to the workbook from which the sheet has been
                       deleted
  @param  ASheetIndex  Index of the sheet that was deleted. The sheet itself
                       does not exist any more.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetRemovedHandler(Sender: TObject;
  ASheetIndex: Integer);
var
  i, sheetCount: Integer;
  sheet: TsWorksheet;
begin
  // It is very possible that the currently selected worksheet has been deleted.
  // Look for the selected worksheet in the workbook. Does it still exist? ...
  i := Workbook.GetWorksheetIndex(FWorksheet);
  if i = -1 then
  begin
    // ... no - it must have been the sheet deleted.
    // We have to select another worksheet.
    sheetCount := Workbook.GetWorksheetCount;
    if (ASheetIndex >= sheetCount) then
      sheet := Workbook.GetWorksheetByIndex(sheetCount-1)
    else
      sheet := Workbook.GetWorksheetbyIndex(ASheetIndex);
  end else
    sheet := FWorksheet;
//  FWorksheet := sheet;  // is needed by listeners!
  NotifyListeners([lniWorksheetRemove]);
  SelectWorksheet(sheet);
end;

{@@ ----------------------------------------------------------------------------
  Event handler called BEFORE a worksheet is deleted.

  @param  Sender      Workbook containing the worksheet
  @param  AWorksheet  Worksheet which is to be deleted
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetRemovingHandler(Sender: TObject;
  AWorksheet: TsWorksheet);
begin
  NotifyListeners([lniWorksheetRemoving], AWorksheet);
end;

{@@ ----------------------------------------------------------------------------
  Event handler called after a worksheet has been renamed

  @param  Sender      Workbook containing the worksheet
  @param  AWorksheet  Worksheet which has been renamed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetRenamedHandler(Sender: TObject;
  AWorksheet: TsWorksheet);
begin
  NotifyListeners([lniWorksheetRename], AWorksheet);
end;

{@@ ----------------------------------------------------------------------------
  Event handler called whenever a the workbook makes a worksheet "active".

  @param  Sender      Workbook containing the worksheet
  @param  AWorksheet  Worksheet which has been activated
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetSelectedHandler(Sender: TObject;
  AWorksheet: TsWorksheet);
var
  r, c: Cardinal;
begin
  FWorksheet := AWorksheet;
  if FWorksheet <> nil then
  begin
    FWorksheet.OnChangeCell := @CellChangedHandler;
    FWorksheet.OnChangeFont := @CellFontChangedHandler;
    FWorksheet.OnSelectCell := @CellSelectedHandler;
    NotifyListeners([lniWorksheet]);
    FWorksheet := AWorksheet;  // !!!!!
    if FWorksheet.ActiveCellRow = Cardinal(-1) then
      r := FWorksheet.TopPaneHeight else
      r := FWorksheet.ActiveCellRow;
    if FWorksheet.ActiveCellCol = Cardinal(-1) then
      c := FWorksheet.LeftPaneWidth else
      c := FWorksheet.ActiveCellCol;
    SelectCell(r, c);
  end else
    NotifyListeners([lniWorksheet]);
end;


{------------------------------------------------------------------------------}
{                            TsWorkbookTabControl                              }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookTabControl.
  Removes itself from the WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsWorkbookTabControl.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  The currently active tab has been changed. The WorkbookSource must activate
  the corresponding worksheet and notify its listening components of the change.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.Change;
begin
  if (FWorkbookSource <> nil) and (FLockCount = 0) then
    FWorkbookSource.SelectWorksheet(Workbook.GetWorksheetByIndex(TabIndex));
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Creates a (string) list containing the names of the workbook's sheet names.
  Is called whenever the workbook changes.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.GetSheetList(AList: TStrings);
var
  i: Integer;
begin
  AList.Clear;
  if Workbook <> nil then
    for i:=0 to Workbook.GetWorksheetCount-1 do
      AList.Add(Workbook.GetWorksheetByIndex(i).Name);
end;

{@@ ----------------------------------------------------------------------------
  Getter method for property "Workbook"
-------------------------------------------------------------------------------}
function TsWorkbookTabControl.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for property "Worksheet"
-------------------------------------------------------------------------------}
function TsWorkbookTabControl.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which
  spreadsheet item has changed.
  Responds to workbook changes by reading the worksheet names into the tabs,
  and to worksheet changes by selecting the tab corresponding to the selected
  worksheet.

  @param  AChangedItems  Set with elements identifying whether workbook, worksheet
                         cell content or cell formatting has changed
  @param  AData          Additional data, not used here
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  i: Integer;
begin
  Unused(AData);

  // Workbook changed: new workbook, worksheet added/renamed/deleted
  if (AChangedItems * [lniWorkbook, lniWorksheetAdd, lniWorksheetRemove, lniWorksheetRename] <> []) then
  begin
    inc(FLockCount);    // avoid WorkbookSelect message when adding each tab
    GetSheetList(Tabs);
    {
    if (lniWorkbook in AChangedItems) then
      TabIndex := 0
    else
    }
    if (lniWorkbook in AChangedItems) and (Workbook <> nil) then
    begin
      i := Workbook.GetWorksheetIndex(Workbook.ActiveWorksheet);
      if i > -1 then TabIndex := i else TabIndex := 0
    end else
    if (lniWorksheetAdd in AChangedItems) then
      TabIndex := Tabs.Count-1
    else
    if (lniWorksheetRename in AChangedItems) then
      TabIndex := Workbook.GetWorksheetIndex(TsWorksheet(AData));
    dec(FLockCount);
  end;

  // Worksheet selected
  if (lniWorksheet in AChangedItems) and (Worksheet <> nil) then
  begin
    i := Tabs.IndexOf(Worksheet.Name);
    if i <> TabIndex then
      TabIndex := i;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification. Must clean up the WorkbookSource field
  when the workbook source is going to be deleted.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the TabControl to the WorkbookSource. Required before
  destruction.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  ListenerNotification([lniWorkbook, lniWorksheet]);
end;


{------------------------------------------------------------------------------}
{                               TsCellEdit                                     }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the spreadsheet edit control. Disables RETURN and TAB keys.
  RETURN characters can still be entered into the edited text by pressing
  CTRL+RETURN

  @param   AOwner   Component which is responsible to destroy the CellEdit
-------------------------------------------------------------------------------}
constructor TsCellEdit.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FShowHTMLText := True;
  WantReturns := false;
  WantTabs := false;
  AutoSize := true;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsCellEdit. Removes itself from the WorkbookSource's
  listener list.
-------------------------------------------------------------------------------}
destructor TsCellEdit.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  EditingDone is called when the user presses the RETURN key to finish editing,
  or the TAB key which removes focus from the control, or clicks somewhere else
  The edited text is written to the worksheet which tries to figure out the
  data type. In particular, if the text begins with an equal sign ("=") then
  the text is assumed to be a formula.
-------------------------------------------------------------------------------}
procedure TsCellEdit.EditingDone;
var
  s: String;
  cell: PCell;
begin
  if Worksheet = nil then
    exit;
  cell := Worksheet.GetCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  if Worksheet.IsMerged(cell) then
    cell := Worksheet.FindMergeBase(cell);
  s := Lines.Text;
  if (s <> '') and (s[1] = '=') then
    Worksheet.WriteFormula(cell, Copy(s, 2, Length(s)), true)
  else
    Worksheet.WriteCellValueAsString(cell, s);
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property SelectedCell which points to the currently
  active cell in the selected worksheet
-------------------------------------------------------------------------------}
function TsCellEdit.GetSelectedCell: PCell;
begin
  if (Worksheet <> nil) then
    with Worksheet do
      Result := FindCell(ActiveCellRow, ActiveCellCol)
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Workbook which is currently loaded by the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellEdit.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Worksheet which is currently "selected" in the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellEdit.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which item
  of the spreadsheet has changed.
  Responds to selection and cell changes by updating the cell content.

  @param  AChangedItems  Set with elements identifying whether workbook, worksheet
                         cell content or cell formatting has changed
  @param  AData          If AChangedItems contains nliCell then AData points to
                         the modified cell.
-------------------------------------------------------------------------------}
procedure TsCellEdit.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  cell: PCell;
begin
  if (FWorkbookSource = nil) then
    exit;

  if  (lniSelection in AChangedItems) or
     ((lniCell in AChangedItems) and (PCell(AData) = SelectedCell))
  then begin
    if Worksheet.IsMerged(SelectedCell) then
      cell := Worksheet.FindMergeBase(SelectedCell)
    else
      cell := SelectedCell;
    ShowCell(cell);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification. Called when the WorkbookSource is deleted.
-------------------------------------------------------------------------------}
procedure TsCellEdit.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the CellEdit to the WorkbookSource. Required before
  destruction.
-------------------------------------------------------------------------------}
procedure TsCellEdit.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsCellEdit.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  Text := '';
  ListenerNotification([lniSelection]);
end;

{@@ ----------------------------------------------------------------------------
  Loads the contents of a cell into the editor.
  Shows the formula if available, but not the calculation result.
  Numbers are displayed in full precision.
  Date and time values are shown in the long formats.

  @param  ACell  Pointer to the cell loaded into the cell editor.
-------------------------------------------------------------------------------}
procedure TsCellEdit.ShowCell(ACell: PCell);
var
  s: String;
begin
  if (FWorkbookSource <> nil) and (ACell <> nil) then
  begin
    s := Worksheet.ReadFormulaAsString(ACell, true);
    if s <> '' then begin
      if s[1] <> '=' then s := '=' + s;
      Lines.Text := s;
    end else
      case ACell^.ContentType of
        cctNumber:
          Lines.Text := FloatToStr(ACell^.NumberValue);
        cctDateTime:
          if ACell^.DateTimeValue < 1.0 then        // Time only
            Lines.Text := FormatDateTime('tt', ACell^.DateTimeValue)
          else
          if frac(ACell^.DateTimeValue) = 0 then    // Date only
            Lines.Text := FormatDateTime('ddddd', ACell^.DateTimevalue)
          else                                      // both
            Lines.Text := FormatDateTime('c', ACell^.DateTimeValue);
        cctUTF8String:
          if FShowHTMLText then
          begin
            RichTextToHTML(Workbook, Worksheet.ReadCellFont(ACell),
              ACell^.UTF8StringValue, ACell^.RichTextParams, s);
            Lines.Text := s;
          end else
            Lines.Text := ACell^.UTF8StringValue;
        else
          Lines.Text := Worksheet.ReadAsUTF8Text(ACell);
      end;
  end else
    Clear;
end;


{------------------------------------------------------------------------------}
{                              TsCellIndicator                                 }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the TsCellIndicator class. Is overridden to set the default
  value of the Alignment property to taCenter.
-------------------------------------------------------------------------------}
constructor TsCellIndicator.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Alignment := taCenter;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the cell indicator. Removes itself from the WorkbookSource's
  listener list.
-------------------------------------------------------------------------------}
destructor TsCellIndicator.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  EditingDone is called when the user presses the RETURN key to finish editing,
  or the TAB key which removes focus from the control, or clicks somewhere else
  The edited text is interpreted as a cell address. The corresponding cell is
  selected.
-------------------------------------------------------------------------------}
procedure TsCellIndicator.EditingDone;
var
  r, c: Cardinal;
begin
  if (WorkbookSource <> nil) and ParseCellString(Text, r, c) then
    WorkbookSource.SelectCell(r, c);
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Workbook which is currently loaded by the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellIndicator.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Worksheet which is currently loaded by the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellIndicator.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  The cell indicator reacts to notification that the selection has changed
  and displays the address of the newly selected cell as editable text.

  @param  AChangedItems  Set with elements identifying whether workbook, worksheet
                         cell or selection has changed. Only the latter element
                         is considered by the cell indicator.
  @param  AData          If AChangedItems contains nliCell then AData points to
                         the modified cell.
-------------------------------------------------------------------------------}
procedure TsCellIndicator.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
var
  sel: TsCellRangeArray;
  s: String;
  rng: TsCellRange;
  numrows, numcols: Integer;
begin
  Unused(AData);
  if (lniSelection in AChangedItems) and (Worksheet <> nil) then
  begin
    s := GetCellString(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
    sel := Worksheet.GetSelection;
    if Length(sel) > 0 then begin
      rng := sel[High(sel)];
      numrows := rng.Row2 - rng.Row1 + 1;
      numcols := rng.Col2 - rng.Col1 + 1;
      if (numrows <> 1) or (numcols <> 1) then
        s := Format('%s (%d R x %d C)', [s, rng.Row2-rng.Row1+1, rng.Col2-rng.Col1+1]);
    end;
    Text := s;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification called when the WorkbookSource is deleted.
-------------------------------------------------------------------------------}
procedure TsCellIndicator.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbooksource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the CellIndicator to the WorkbookSource. Required before
  destruction.
-------------------------------------------------------------------------------}
procedure TsCellIndicator.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsCellIndicator.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  Text := '';
  ListenerNotification([lniSelection]);
end;


{------------------------------------------------------------------------------}
{                               TsCellCombobox                                 }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the Cell Combobox. Populates the items list
-------------------------------------------------------------------------------}
constructor TsCellCombobox.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FColorRectWidth := 10;
  FColorRectOffset := 2;
  ItemHeight := -1;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookTabControl.
  Removes itself from the WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsCellCombobox.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Applies the format to a cell. Override according to the format item for
  which the combobox is responsible.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ApplyFormatToCell(ACell: PCell);
var
  fnt: TsFont;
  clr: TColor;
begin
  if (Worksheet = nil) then
    exit;

  case FFormatItem of
    cfiFontName:
      if Text <> '' then
      begin
        fnt := Worksheet.ReadCellFont(ACell);
        Worksheet.WriteFont(ACell, Text, fnt.Size, fnt.Style, fnt.Color);
      end;
    cfiFontSize:
      if Text <> '' then
      begin
        fnt := Worksheet.ReadCellFont(ACell);
        Worksheet.WriteFont(ACell, fnt.FontName, StrToFloat(Text), fnt.Style, fnt.Color);
      end;
    cfiFontColor:
      if ItemIndex > -1 then
      begin
        fnt := Worksheet.ReadCellFont(ACell);
        clr := PtrInt(Items.Objects[ItemIndex]);
        Worksheet.WriteFont(ACell, fnt.FontName, fnt.Size, fnt.style, clr);
      end;
    cfiBackgroundColor:
      if ItemIndex <= 0 then
        Worksheet.WriteBackgroundColor(ACell, scTransparent)
      else
      begin
        clr := PtrInt(Items.Objects[ItemIndex]);
        Worksheet.WriteBackgroundColor(ACell, clr);
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatCombobox.ApplyFormatToCell] Unknown format item');
  end;
end;

{@@ ----------------------------------------------------------------------------
  The text of the currently selected combobox item has been changed.
  Calls "ProcessValue" to changes the selected cells according to the
  Mode property by calling ApplyFormatToCell.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.Change;
begin
  inherited;
  ProcessItem;
end;

{@@ ----------------------------------------------------------------------------
  Customdraws an item in the combobox. This is overridden to paint a color box
  for the color-related format items.
 ------------------------------------------------------------------------------}
procedure TsCellCombobox.DrawItem(AIndex: Integer; ARect: TRect;
  AState: TOwnerDrawState);
{ This code is adapted from colorbox.pas}
var
  r: TRect;
  clr: TsColor;
  brushColor, penColor: TColor;
  brushStyle: TBrushStyle;
  noFill: Boolean;
begin
  if AIndex = -1 then
    Exit;

  if FFormatItem in [cfiFontColor, cfiBackgroundColor, cfiBorderColor] then
  begin
    r.Top := ARect.Top + FColorRectOffset;
    r.Bottom := ARect.Bottom - FColorRectOffset;
    r.Left := ARect.Left + FColorRectOffset;
    if FColorRectWidth = -1 then
      r.Right := ARect.Right - FColorRectOffset
    else
      r.Right := r.Left + FColorRectWidth;
    Exclude(AState, odPainted);

    noFill := false;

    with Canvas do
    begin
      FillRect(ARect);

      brushStyle := Brush.Style;
      brushColor := Brush.Color;
      penColor := Pen.Color;

      clr := TsColor(PtrInt(Items.Objects[AIndex]));
      if (clr = scTransparent) or (clr = scNotDefined) then
      begin
        noFill := true;
        Brush.Style := bsClear;
      end else
      begin
        Brush.Color := clr and $00FFFFFF;
        Brush.Style := bsSolid;
      end;
      Pen.Color := clBlack;

      r := BiDiFlipRect(r, ARect, UseRightToLeftAlignment);
      Rectangle(r);

      if noFill then
      begin
        Line(r.Left, r.Top, r.Right-1, r.Bottom-1);
        Line(r.Left, r.Bottom-1, r.Right-1, r.Top);
      end;

      Brush.Style := brushStyle;
      Brush.Color := brushColor;
      Pen.Color := penColor;
    end;

    if FColorRectWidth > -1 then
    begin
      r := ARect;
      inc(r.Left, FColorRectWidth + 2*FColorRectOffset);
      inherited DrawItem(AIndex, BidiFlipRect(r, ARect, UseRightToLeftAlignment), AState);
    end;
  end else
  begin
    r := ARect;
    inherited DrawItem(AIndex, BidiFlipRect(r, ARect, UseRightToLeftAlignment), AState);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Extracts the format item the combobox is responsible for from the cell and
  selectes the corresponding combobox item.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ExtractFromCell(ACell: PCell);
var
  fnt: TsFont;
  clr: TsColor;
begin
  case FFormatItem of
    cfiFontName:
      begin
        fnt := Worksheet.ReadCellFont(ACell);
        // No check for nil required because fnt is at least DefaultFont
        ItemIndex := Items.IndexOf(fnt.FontName);
      end;
    cfiFontSize:
      begin
        fnt := Worksheet.ReadCellFont(ACell);
        ItemIndex := Items.IndexOf(Format('%.0f', [fnt.Size]));
      end;
    cfiFontColor:
      begin
        fnt := Worksheet.ReadCellFont(ACell);
        ItemIndex := Items.IndexOfObject(TObject(PtrInt(fnt.Color)));
      end;
    cfiBackgroundColor:
      begin
        clr := Worksheet.ReadBackgroundColor(ACell);
        ItemIndex := Max(0, Items.IndexOfObject(TObject(PtrInt(clr))));
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatItem.ExtractFromCell] Unknown format item');
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the currently active cell of the worksheet
-------------------------------------------------------------------------------}
function TsCellCombobox.GetActiveCell: PCell;
begin
  if FWorkbookSource <> nil then
    Result := Worksheet.FindCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol)
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Workbook which is currently loaded by the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellCombobox.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Worksheet which is currently loaded by the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsCellCombobox.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Notification procedure received whenver "something" changes in the workbook.
  Reacts on all events.

  @param  AChangedItems  Set with elements identifying whether workbook, worksheet
                         cell or selection has changed.
  @param  AData          If AChangedItems contains nliCell then AData points to
                         the modified cell.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  activeCell: PCell;
begin
  Unused(AData);
  if (Worksheet = nil) or
     ([lniCell, lniSelection]*AChangedItems = [])
  then
    exit;

  activeCell := GetActiveCell;
  if (([lniCell]*AChangedItems <> []) and (PCell(AData) = activeCell)) or
     (lniSelection in AChangedItems)
  then
    ExtractFromCell(activeCell);
                                  {
  if (FFormatItem in [cfiFontColor, cfiBorderColor, cfiBackgroundColor]) then
    Populate;                      }
end;

{@@ ----------------------------------------------------------------------------
  Standard method. Overridden to populate combobox since items are not stored
  in lfm file.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.Loaded;
begin
  inherited;
  Populate;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification method called when the WorkbookSource
  is deleted.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Descendants override this method to populate the items of the combobox.
  Color index into the workbook's palette is stored in the "Objects" property.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.Populate;
var
  i: Integer;
  clr: TsColor;
  noText: Boolean;
begin
  if Workbook = nil then
    exit;

  case FFormatItem of
    cfiFontName:
      Items.Assign(Screen.Fonts);
    cfiFontSize:
      Items.CommaText := '8,9,10,11,12,13,14,16,18,20,22,24,26,28,32,36,48,72';
    cfiBackgroundColor,
    cfiFontColor,
    cfiBorderColor:
      begin
        noText := (FColorRectWidth = -1);
        Items.Clear;
        if FFormatItem = cfiBackgroundColor then
          Items.AddObject(StrUtils.IfThen(noText, '', '(none)'), TObject(scTransparent));
        for i:=0 to ComboColors.Count-1 do
        begin
          clr := ComboColors[i];
          Items.AddObject(StrUtils.IfThen(noText, '', GetColorName(clr)), TObject(PtrInt(clr)));
        end;
      end;
    else
      raise Exception.Create('[TsCellCombobox.Populate] Unknown cell format item.');
  end;
end;

{@@ ----------------------------------------------------------------------------
  Processes the selected combobox item after a new item has been selected or the
  item text has been edited.
  Changes the selected cells according to the Mode property by calling
  ApplyFormatToCell.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ProcessItem;
var
  r, c: Cardinal;
  range: Integer;
  sel: TsCellRangeArray;
  cell: PCell;
begin
  if Worksheet = nil then
    exit;
  sel := Worksheet.GetSelection;
  if Length(sel) = 0 then
  begin
    SetLength(sel, 1);
    sel[0].Col1 := Worksheet.ActiveCellCol;
    sel[0].Row1 := Worksheet.ActiveCellRow;
    sel[0].Col2 := sel[0].Col1;
    sel[0].Row2 := sel[0].Row2;
  end;
  for range := 0 to High(sel) do
    for r := sel[range].Row1 to sel[range].Row2 do
      for c := sel[range].Col1 to sel[range].Col2 do
      begin
        cell := Worksheet.GetCell(r, c);  // Use "GetCell" here to format empty cells as well
        ApplyFormatToCell(cell);  // no check for nil required because of "GetCell"
      end;
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the CellCombobox to the WorkbookSource. Required before
  destruction.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  A new item in the combobox is selected. Calls "ProcessValue" to changes the
  selected cells according to the Mode property by calling ApplyFormatToCell.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.Select;
begin
  inherited Select;
  ProcessItem;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the ColorRectOffset property
-------------------------------------------------------------------------------}
procedure TsCellCombobox.SetColorRectOffset(AValue: Integer);
begin
  if FColorRectOffset = AValue then
    exit;
  FColorRectOffset := AValue;
  Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the ColorRectWidth property
-------------------------------------------------------------------------------}
procedure TsCellCombobox.SetColorRectWidth(AValue: Integer);
begin
  if FColorRectWidth = AValue then
    exit;
  FColorRectWidth := AValue;
  Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the FormatItem property
-------------------------------------------------------------------------------}
procedure TsCellCombobox.SetFormatItem(AValue: TsCellFormatItem);
begin
  FFormatItem := AValue;
  if FFormatItem in [cfiFontColor, cfiBackgroundColor, cfiBorderColor] then
  begin
    inherited Style := csOwnerDrawFixed;
    ReadOnly := true;
  end else
  begin
    inherited Style := csDropdown;
    ReadOnly := false;
  end;

  Populate;
  if FWorkbookSource <> nil then
    ExtractFromCell(GetActiveCell);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsCellCombobox.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  Text := '';
  ListenerNotification([lniSelection]);
end;
  (*
procedure TsCellCombobox.UpdateCombo;
var
  c: integer;
begin
  if HandleAllocated then
    Invalidate;
  {
  begin
    for c := Ord(cbCustomColor in Style) to Items.Count - 1 do
    begin
      if Colors[c] = FSelected then
      begin
        ItemIndex := c;
        Exit;
      end;
    end;
    if cbCustomColor in Style then
    begin
      Items.Objects[0] := TObject(PtrInt(FSelected));
      ItemIndex := 0;
      Invalidate;
    end
    else
      ItemIndex := -1;
  end;
  }
end;
    *)
                 (*
{------------------------------------------------------------------------------}
{                             TsCellFontCombobox                               }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Determines the font used by a specified cell. Returns the workbook's default
  font if the cell does not exist. Considers the uffBold and uffFont formatting
  fields of the cell
-------------------------------------------------------------------------------}
function TsCellFontCombobox.GetCellFont(ACell: PCell): TsFont;
begin
  if ACell = nil then
    Result := Workbook.GetDefaultFont
  else
  if (uffBold in ACell^.UsedFormattingFields) then
    Result := Workbook.GetFont(1)
  else
  if (uffFont in ACell^.UsedFormattingFields) then
    Result := Workbook.GetFont(ACell^.FontIndex)
  else
    Result := Workbook.GetDefaultFont;
end;


{------------------------------------------------------------------------------}
{                             TsFontNameCombobox                               }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the FontNameCombobox. Predefines the width of the combobox
  such that it is sufficient for most font names
-------------------------------------------------------------------------------}
constructor TsFontNameCombobox.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Width := 150;
end;

{@@ ----------------------------------------------------------------------------
  Applies the font with the selected name to a specified cell.
-------------------------------------------------------------------------------}
procedure TsFontNameCombobox.ApplyFormatToCell(ACell: PCell);
var
  fnt: TsFont;
begin
  if ItemIndex > -1 then
  begin
    fnt := GetCellFont(ACell);
    Worksheet.WriteFont(ACell, Items[ItemIndex], fnt.Size, fnt.Style, fnt.Color);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Extracts the font of the specified cell and selects its font name in the
  combobox.
-------------------------------------------------------------------------------}
procedure TsFontNameCombobox.ExtractFromCell(ACell: PCell);
var
  fnt: TsFont;
begin
  fnt := GetCellFont(ACell);
  if fnt <> nil then
    ItemIndex := Items.IndexOf(fnt.FontName);
end;

{@@ ----------------------------------------------------------------------------
  Populates the combobox with the names of all fonts available on the current
  system
-------------------------------------------------------------------------------}
procedure TsFontNameCombobox.Populate;
begin
  Items.Assign(Screen.Fonts);
end;


{------------------------------------------------------------------------------}
{                             TsFontSizeCombobox                               }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the FontSizeCombobox. Reduces the default width of the combobox
  due to the narrow width of the font size numbers.
-------------------------------------------------------------------------------}
constructor TsFontSizeCombobox.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Width := 60;
end;

{@@ ----------------------------------------------------------------------------
  Applies the font with the selected size to a specified cell.
-------------------------------------------------------------------------------}
procedure TsFontSizeCombobox.ApplyFormatToCell(ACell: PCell);
var
  fnt: TsFont;
  fs: Double;
begin
  if ItemIndex > -1 then
  begin
    fs := StrToFloat(Items[ItemIndex]);
    fnt := GetCellFont(ACell);
    Worksheet.WriteFont(ACell, fnt.FontName, fs, fnt.Style, fnt.Color);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Extracts the font of the specified cell and selects its font size in the
  combobox.
-------------------------------------------------------------------------------}
procedure TsFontSizeCombobox.ExtractFromCell(ACell: PCell);
var
  fnt: TsFont;
begin
  fnt := GetCellFont(ACell);
  if fnt <> nil then
    ItemIndex := Items.IndexOf(Format('%.0f', [fnt.Size]));
end;

{@@ ----------------------------------------------------------------------------
  Populates the combobox with often-used font sizes (in points)
-------------------------------------------------------------------------------}
procedure TsFontSizeCombobox.Populate;
begin
  with Items do
  begin
    Clear;
    Add('8');
    Add('9');
    Add('10');
    Add('11');
    Add('12');
    Add('14');
    Add('16');
    Add('18');
    Add('20');
    Add('22');
    Add('24');
    Add('26');
    Add('28');
    Add('32');
    Add('36');
    Add('48');
    Add('72');
  end;
end;
            *)

{------------------------------------------------------------------------------}
{                          TsSpreadsheetInspector                              }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the TsSpreadsheetInspector class.
  Is overridden to set the default values of DisplayOptions and FixedCols, and
  to define the column captions.
-------------------------------------------------------------------------------}
constructor TsSpreadsheetInspector.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DisplayOptions := DisplayOptions - [doKeyColFixed];
  FixedCols := 0;
  FExpanded := [ienFormatSettings, ienPageLayout];
  with (TitleCaptions as TStringList) do begin
    OnChange := nil;        // This fixes an issue with Laz 1.0
    Clear;
    Add('Properties');
    Add('Values');
    OnChange := @TitlesChanged;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the spreadsheet inspector. Removes itself from the
  WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsSpreadsheetInspector.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Double-click on an expandable line expands or collapsed the sub-items
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.DblClick;
var
  s: String;
  expNodes: TsInspectorExpandedNodes;
begin
  expNodes := FExpanded;
  s := Cells[0, Row];
  if (pos('FormatSettings', s) > 0) or (pos('Format settings', s) > 0) then
  begin
    if (ienFormatSettings in expNodes)
      then Exclude(expNodes, ienFormatSettings)
      else Include(expNodes, ienFormatSettings);
  end else
  if (pos('Page layout', s) > 0) or (pos('PageLayout', s) > 0) then
  begin
    if (ienPageLayout in expNodes)
      then Exclude(expNodes, ienPageLayout)
      else Include(expNodes, ienPageLayout);
  end else
    exit;
  SetExpanded(expNodes);
end;

{@@ ----------------------------------------------------------------------------
  Updates the data shown by the inspector grid. Display depends on the FMode
  setting (workbook, worksheet, cell values, cell properties).
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.DoUpdate;
var
  cell: PCell;
  sheet: TsWorksheet;
  book: TsWorkbook;
  list: TStringList;
begin
  cell := nil;
  sheet := nil;
  book := nil;
  if FWorkbookSource <> nil then
  begin
    book := FWorkbookSource.Workbook;
    sheet := FWorkbookSource.Worksheet;
    if sheet <> nil then
      cell := sheet.FindCell(sheet.ActiveCellRow, sheet.ActiveCellCol);
  end;

  list := TStringList.Create;
  try
    case FMode of
      imCellValue      : UpdateCellValue(cell, list);
      imCellProperties : UpdateCellProperties(cell, list);
      imWorksheet      : UpdateWorksheet(sheet, list);
      imWorkbook       : UpdateWorkbook(book, list);
    end;
    Strings.Assign(list);
  finally
    list.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Workbook which is currently loaded by the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsSpreadsheetInspector.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the property Worksheet which is currently loaded by the
  WorkbookSource
-------------------------------------------------------------------------------}
function TsSpreadsheetInspector.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Notification procedure received whenver "something" changes in the workbook.
  Reacts on all events.

  @param  AChangedItems  Set with elements identifying whether workbook, worksheet
                         cell or selection has changed.
  @param  AData          If AChangedItems contains nliCell then AData points to
                         the modified cell.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
begin
  Unused(AData);
  case FMode of
    imWorkbook:
      if ([lniWorkbook, lniWorksheet]*AChangedItems <> []) then DoUpdate;
    imWorksheet:
      if ([lniWorksheet, lniSelection]*AChangedItems <> []) then DoUpdate;
    imCellValue, imCellProperties:
      if ([lniCell, lniSelection]*AChangedItems <> []) then DoUpdate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification method called when the WorkbookSource
  is deleted.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the SpreadsheetInspector to the WorkbookSource.
  Required before destruction.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the Expanded property
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.SetExpanded(AValue: TsInspectorExpandedNodes);
begin
  if AValue = FExpanded then
    exit;
  FExpanded := AValue;
  DoUpdate;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the Mode property. This property filters groups of properties
  for display (workbook-, worksheet-, cell value- or cell formatting-related
  data).
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.SetMode(AValue: TsInspectorMode);
begin
  if AValue = FMode then
    exit;
  FMode := AValue;
  DoUpdate;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  ListenerNotification([lniWorkbook, lniWorksheet, lniSelection]);
end;

{@@ ----------------------------------------------------------------------------
  Creates a string list containing the formatting properties of a specific cell.
  The string list items are name-value pairs in the format "name=value".
  The string list is displayed in the inspector's grid.

  @param  ACell     Pointer to cell under investigation
  @param  AStrings  Stringlist receiving the name-value pairs.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.UpdateCellProperties(ACell: PCell;
  AStrings: TStrings);
var
  s: String;
  cb: TsCellBorder;
  r1, r2, c1, c2: Cardinal;
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams;
  rtp: TsRichTextParam;
begin
  if (ACell <> nil) then
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex)
  else
    InitFormatRecord(fmt);

  if (ACell = nil)
    then AStrings.Add('FormatIndex=(default)')
    else AStrings.Add(Format('FormatIndex=%d', [ACell^.FormatIndex]));

  if (ACell = nil) or not (uffFont in fmt.UsedFormattingFields)
    then AStrings.Add('FontIndex=(default)')
    else AStrings.Add(Format('FontIndex=%d (%s)', [
           fmt.FontIndex,
           Workbook.GetFontAsString(fmt.FontIndex)
         ]));

  if (ACell <> nil) and (Length(ACell^.RichTextParams) > 0) then
  begin
    s := '';
    for rtp in ACell^.RichTextParams do
      s := Format('%s; Font #%d after pos %d', [s, rtp.FontIndex, rtp.FirstIndex]);
    Delete(s, 1, 2);
    if s = '' then s := '(none)';
    AStrings.Add('Rich-text parameters='+s);
  end else
    AStrings.Add('Rich-text parameters=(none)');

  if (ACell=nil) or not (uffTextRotation in fmt.UsedFormattingFields)
    then AStrings.Add('TextRotation=(default)')
    else AStrings.Add(Format('TextRotation=%s', [
           GetEnumName(TypeInfo(TsTextRotation), ord(fmt.TextRotation))
         ]));

  if (ACell=nil) or not (uffHorAlign in fmt.UsedFormattingFields)
    then AStrings.Add('HorAlignment=(default)')
    else AStrings.Add(Format('HorAlignment=%s', [
           GetEnumName(TypeInfo(TsHorAlignment), ord(fmt.HorAlignment))
         ]));

  if (ACell=nil) or not (uffVertAlign in fmt.UsedFormattingFields)
    then AStrings.Add('VertAlignment=(default)')
    else AStrings.Add(Format('VertAlignment=%s', [
           GetEnumName(TypeInfo(TsVertAlignment), ord(fmt.VertAlignment))
         ]));

  if (ACell=nil) or not (uffBorder in fmt.UsedFormattingFields) then
    AStrings.Add('Borders=(none)')
  else
  begin
    s := '';
    for cb in TsCellBorder do
      if cb in fmt.Border then
        s := s + ', ' + GetEnumName(TypeInfo(TsCellBorder), ord(cb));
    if s <> '' then Delete(s, 1, 2);
    AStrings.Add('Borders='+s);
  end;

  for cb in TsCellBorder do
    if ACell = nil then
      AStrings.Add(Format('BorderStyles[%s]=(default)', [
        GetEnumName(TypeInfo(TsCellBorder), ord(cb))]))
    else
      AStrings.Add(Format('BorderStyles[%s]=%s, %s', [
        GetEnumName(TypeInfo(TsCellBorder), ord(cb)),
        GetEnumName(TypeInfo(TsLineStyle), ord(fmt.BorderStyles[cb].LineStyle)),
        GetColorName(fmt.BorderStyles[cb].Color)]));

  if (ACell = nil) or not (uffBackground in fmt.UsedformattingFields) then
  begin
    AStrings.Add('Style=(default)');
    AStrings.Add('PatternColor=(default)');
    AStrings.Add('BackgroundColor=(default)');
  end else
  begin
    AStrings.Add(Format('Style=%s', [
      GetEnumName(TypeInfo(TsFillStyle), ord(fmt.Background.Style))]));
    AStrings.Add(Format('PatternColor=$%.8x (%s)', [
      fmt.Background.FgColor, GetColorName(fmt.Background.FgColor)]));
    AStrings.Add(Format('BackgroundColor=$%.8x (%s)', [
      fmt.Background.BgColor, GetColorName(fmt.Background.BgColor)]));
  end;

  if (ACell = nil) or not (uffNumberFormat in fmt.UsedFormattingFields) then
  begin
    AStrings.Add('NumberFormatIndex=-1');
    AStrings.Add('NumberFormat=(default)');
    AStrings.Add('NumberFormatStr=(none)');
  end else
  begin
    AStrings.Add(Format('NumberFormatIndex=%d', [fmt.NumberFormatIndex]));
    numFmt := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
    AStrings.Add(Format('NumberFormat=%s', [
      GetEnumName(TypeInfo(TsNumberFormat), ord(numFmt.NumFormat))]));
    AStrings.Add('NumberFormatStr=' + numFmt.NumFormatStr);
  end;

  if (Worksheet = nil) or not Worksheet.IsMerged(ACell) then
    AStrings.Add('Merged range=(none)')
  else
  begin
    Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    AStrings.Add('Merged range=' + GetCellRangeString(r1, c1, r2, c2));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a string list containing the value data of a specific cell.
  The string list items are name-value pairs in the format "name=value".
  The string list is displayed in the inspector's grid.

  @param  ACell     Pointer to cell under investigation
  @param  AStrings  Stringlist receiving the name-value pairs.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.UpdateCellValue(ACell: PCell; AStrings: TStrings);
var
  hyperlink: PsHyperlink;
  comment: String;
  s: String;
begin
  if ACell = nil then
  begin
    if Worksheet <> nil then
    begin
      AStrings.Add(Format('Row=%d', [Worksheet.ActiveCellRow]));
      AStrings.Add(Format('Col=%d', [Worksheet.ActiveCellCol]));
    end else
    begin
      AStrings.Add('Row=');
      AStrings.Add('Col=');
    end;
    AStrings.Add('ContentType=(none)');
  end else
  begin
    AStrings.Add(Format('Row=%d', [ACell^.Row]));
    AStrings.Add(Format('Col=%d', [ACell^.Col]));
    AStrings.Add(Format('Flags=[%s]', [
      SetToString(PTypeInfo(TypeInfo(TsCellflags)), integer(ACell^.Flags), false)
    ]));
    AStrings.Add(Format('ContentType=%s', [
      GetEnumName(TypeInfo(TCellContentType), ord(ACell^.ContentType))
    ]));
    if ACell^.ContentType = cctNumber then
      AStrings.Add(Format('NumberValue=%g', [ACell^.NumberValue]));
    if ACell^.ContentType = cctDateTime then
      AStrings.Add(Format('DateTimeValue=%g', [ACell^.DateTimeValue]));
    if ACell^.ContentType = cctUTF8String then
      AStrings.Add(Format('UTF8StringValue=%s', [ACell^.UTF8StringValue]));
    if ACell^.ContentType = cctBool then
      AStrings.Add(Format('BoolValue=%s', [BoolToStr(ACell^.BoolValue)]));
    if ACell^.ContentType = cctError then
      AStrings.Add(Format('ErrorValue=%s', [GetEnumName(TypeInfo(TsErrorValue), ord(ACell^.ErrorValue))]));
    AStrings.Add(Format('FormulaValue=%s', [Worksheet.ReadFormulaAsString(ACell, true)]));
    {
    if ACell^.SharedFormulaBase = nil then
      AStrings.Add('SharedFormulaBase=')
    else
      AStrings.Add(Format('SharedFormulaBase=%s', [GetCellString(
        ACell^.SharedFormulaBase^.Row, ACell^.SharedFormulaBase^.Col)
      ]));
      }
    if (cfHyperlink in ACell^.Flags) then
    begin
      hyperlink := Worksheet.FindHyperlink(ACell);
      if hyperlink <> nil then
      begin
        if hyperlink^.Tooltip <> '' then
          s := hyperlink^.Target + ' (tooltip: ' + hyperlink^.Tooltip + ')'
        else
          s := hyperlink^.Target;
        AStrings.Add(Format('Hyperlink=%s', [s]));
      end;
    end;
    if (cfHasComment in ACell^.Flags) then
    begin
      comment := Worksheet.ReadComment(ACell);
      AStrings.Add(Format('Comment=%s', [comment]));
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a string list containing the properties of the workbook.
  The string list items are name-value pairs in the format "name=value".
  The string list is displayed in the inspector's grid.

  @param  AWorkbook  Workbook under investigation
  @param  AStrings   Stringlist receiving the name-value pairs.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.UpdateWorkbook(AWorkbook: TsWorkbook;
  AStrings: TStrings);
var
  bo: TsWorkbookOption;
  s: String;
  i: Integer;
begin
  if AWorkbook = nil then
  begin
    AStrings.Add('FileName=');
    AStrings.Add('FileFormat=');
    AStrings.Add('Options=');
    AStrings.Add('ActiveWorksheet=');
    AStrings.Add('FormatSettings=');
  end else
  begin
    AStrings.Add(Format('FileName=%s', [AWorkbook.FileName]));
    AStrings.Add(Format('FileFormat=%d [%s]', [
      AWorkbook.FileFormatID, GetSpreadTechnicalName(AWorkbook.FileFormatID)
    ]));
    {
    AStrings.Add(Format('FileFormat=%s', [
      GetEnumName(TypeInfo(TsSpreadsheetFormat), ord(AWorkbook.FileFormat))
    ]));
     }
    if AWorkbook.ActiveWorksheet <> nil then
      AStrings.Add('ActiveWorksheet=' + AWorkbook.ActiveWorksheet.Name)
    else
      AStrings.Add('ActiveWorksheet=');

    s := '';
    for bo in TsWorkbookOption do
      if bo in AWorkbook.Options then
        s := s + ', ' + GetEnumName(TypeInfo(TsWorkbookOption), ord(bo));
    if s <> '' then Delete(s, 1, 2);
    AStrings.Add('Options='+s);

    if (ienFormatSettings in FExpanded) then begin
      AStrings.Add('(-) FormatSettings=');
      AStrings.Add('  ThousandSeparator='+AWorkbook.FormatSettings.ThousandSeparator);
      AStrings.Add('  DecimalSeparator='+AWorkbook.FormatSettings.DecimalSeparator);
      AStrings.Add('  ListSeparator='+AWorkbook.FormatSettings.ListSeparator);
      AStrings.Add('  DateSeparator='+AWorkbook.FormatSettings.DateSeparator);
      AStrings.Add('  TimeSeparator='+AWorkbook.FormatSettings.TimeSeparator);
      AStrings.Add('  ShortDateFormat='+AWorkbook.FormatSettings.ShortDateFormat);
      AStrings.Add('  LongDateFormat='+AWorkbook.FormatSettings.LongDateFormat);
      AStrings.Add('  ShortTimeFormat='+AWorkbook.FormatSettings.ShortTimeFormat);
      AStrings.Add('  LongTimeFormat='+AWorkbook.FormatSettings.LongTimeFormat);
      AStrings.Add('  TimeAMString='+AWorkbook.FormatSettings.TimeAMString);
      AStrings.Add('  TimePMString='+AWorkbook.FormatSettings.TimePMString);
      s := AWorkbook.FormatSettings.ShortMonthNames[1];
      for i:=2 to 12 do
        s := s + ', ' + AWorkbook.FormatSettings.ShortMonthNames[i];
      AStrings.Add('  ShortMonthNames='+s);
      s := AWorkbook.FormatSettings.LongMonthnames[1];
      for i:=2 to 12 do
        s := s +', ' + AWorkbook.FormatSettings.LongMonthNames[i];
      AStrings.Add('  LongMontNames='+s);
      s := AWorkbook.FormatSettings.ShortDayNames[1];
      for i:=2 to 7 do
        s := s + ', ' + AWorkbook.FormatSettings.ShortDayNames[i];
      AStrings.Add('  ShortMonthNames='+s);
      s := AWorkbook.FormatSettings.LongDayNames[1];
      for i:=2 to 7 do
        s := s +', ' + AWorkbook.FormatSettings.LongDayNames[i];
      AStrings.Add('  LongMontNames='+s);
      AStrings.Add('  CurrencyString='+AWorkbook.FormatSettings.CurrencyString);
      AStrings.Add('  PosCurrencyFormat='+IntToStr(AWorkbook.FormatSettings.CurrencyFormat));
      AStrings.Add('  NegCurrencyFormat='+IntToStr(AWorkbook.FormatSettings.NegCurrFormat));
      AStrings.Add('  TwoDigitYearCenturyWindow='+IntToStr(AWorkbook.FormatSettings.TwoDigitYearCenturyWindow));
    end else
      AStrings.Add('(+) FormatSettings=(dblclick for more...)');

    for i:=0 to AWorkbook.GetFontCount-1 do
      AStrings.Add(Format('Font%d=%s', [i, AWorkbook.GetFontAsString(i)]));

    for i:=0 to AWorkbook.GetNumCellFormats-1 do
      AStrings.Add(Format('CellFormat%d=%s', [i, AWorkbook.GetCellFormatAsString(i)]));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a string list containing the properties of a worksheet.
  The string list items are name-value pairs in the format "name=value".
  The string list is displayed in the inspector's grid.

  @param  ASheet    Worksheet under investigation
  @param  AStrings  Stringlist receiving the name-value pairs.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.UpdateWorksheet(ASheet: TsWorksheet;
  AStrings: TStrings);
var
  s: String;
  po: TsPrintOption;
begin
  if ASheet = nil then
  begin
    AStrings.Add('Name=');
    AStrings.Add('First row=');
    AStrings.Add('Last row=');
    AStrings.Add('First column=');
    AStrings.Add('Last column=');
    AStrings.Add('Active cell=');
    AStrings.Add('Selection=');
    AStrings.Add('Page layout=');
  end else
  begin
    AStrings.Add(Format('Name=%s', [ASheet.Name]));
    AStrings.Add(Format('First row=%d', [Integer(ASheet.GetFirstRowIndex)]));
    AStrings.Add(Format('Last row=%d', [ASheet.GetLastRowIndex]));
    AStrings.Add(Format('First column=%d', [Integer(ASheet.GetFirstColIndex)]));
    AStrings.Add(Format('Last column=%d', [ASheet.GetLastColIndex]));
    AStrings.Add(Format('Active cell=%s', [GetCellString(ASheet.ActiveCellRow, ASheet.ActiveCellCol)]));
    AStrings.Add(Format('Selection=%s', [ASheet.GetSelectionAsString]));
    AStrings.Add(Format('Comments=%d items', [ASheet.Comments.Count]));
    AStrings.Add(Format('Hyperlinks=%d items', [ASheet.Hyperlinks.Count]));
    AStrings.Add(Format('MergedCells=%d items', [ASheet.MergedCells.Count]));
    if ienPageLayout in FExpanded then
    begin
      AStrings.Add('(+) Page layout=');
      AStrings.Add(Format('  Orientation=%s', [GetEnumName(TypeInfo(TsPageOrientation), ord(ASheet.PageLayout.Orientation))]));
      AStrings.Add(Format('  Page width=%.1f mm', [ASheet.PageLayout.PageWidth]));
      AStrings.Add(Format('  Page height=%.1f mm', [ASheet.PageLayout.PageHeight]));
      AStrings.Add(Format('  Left margin=%.1f mm', [ASheet.PageLayout.LeftMargin]));
      AStrings.Add(Format('  Right margin=%.1f mm', [ASheet.PageLayout.RightMargin]));
      AStrings.Add(Format('  Top margin=%.1f mm', [ASheet.PageLayout.TopMargin]));
      AStrings.Add(Format('  Bottom margin=%.1f mm', [ASheet.PageLayout.BottomMargin]));
      AStrings.Add(Format('  Header distance=%.1f mm', [ASheet.PageLayout.HeaderMargin]));
      AStrings.Add(Format('  Footer distance=%.1f mm', [ASheet.PageLayout.FooterMargin]));
      if poUseStartPageNumber in ASheet.PageLayout.Options then
        AStrings.Add(Format('  Start page number=%d', [ASheet.pageLayout.StartPageNumber]))
      else
        AStrings.Add     ('  Start page number=automatic');
      AStrings.Add(Format('  Scaling factor=%d%%', [ASheet.PageLayout.ScalingFactor]));
      AStrings.Add(Format('  Copies=%d', [ASheet.PageLayout.Copies]));
      if (ASheet.PageLayout.Options * [poDifferentOddEven, poDifferentFirst] <> []) then
      begin
        AStrings.Add(Format('  Header (first)=%s', [StringReplace(ASheet.PageLayout.Headers[0], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Header (odd)=%s', [StringReplace(ASheet.PageLayout.Headers[1], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Header (even)=%s', [StringReplace(ASheet.PageLayout.Headers[2], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Footer (first)=%s', [StringReplace(ASheet.PageLayout.Footers[0], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Footer (odd)=%s', [StringReplace(ASheet.PageLayout.Footers[1], LineEnding, '\n', [rfReplaceall])]));
        AStrings.Add(Format('  Footer (even)=%s', [StringReplace(ASheet.PageLayout.Footers[2], LineEnding, '\n', [rfReplaceAll])]));
      end else
      begin
        AStrings.Add(Format('  Header=%s', [StringReplace(ASheet.PageLayout.Headers[1], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Footer=%s', [StringReplace(ASheet.PageLayout.Footers[1], LineEnding, '\n', [rfReplaceAll])]));
      end;
      s := '';
      for po in TsPrintOption do
        if po in ASheet.PageLayout.Options then s := s + '; ' + GetEnumName(typeInfo(TsPrintOption), ord(po));
      if s <> '' then Delete(s, 1, 2);
      AStrings.Add(Format('  Options=%s', [s]));
    end else
      AStrings.Add('(+) Page layout=(dblclick for more...)');
  end;
end;

initialization
  {$I fpspreadsheetctrls.lrs}

//  CellClipboard := TsCellList.Create;

  ComboColors := TsPalette.Create;
  ComboColors.AddExcelColors;

  RegisterPropertyToSkip(TsSpreadsheetInspector, 'RowHeights', 'For compatibility with older Laz versions.', '');
  RegisterPropertyToSkip(TsSpreadsheetInspector, 'ColWidths', 'For compatibility with older Laz versions.', '');

  { Clipboard formats }
  cfBiff8Format := RegisterclipboardFormat('Biff8');
  cfBiff5Format := RegisterClipboardFormat('Biff5');
  cfHTMLFormat := RegisterClipboardFormat('HTML Format');
  cfTextHTMLFormat := RegisterClipboardFormat('text/html');
  cfCSVFormat := RegisterClipboardFormat('CSV');
  { not working...
  cfOpenDocumentFormat := RegisterClipboardFormat('application/x-openoffice-embed-source-xml;windows_formatname="Star Embed Source (XML)"');
  cfStarObjectDescriptor := RegisterClipboardFormat('application/x-openoffice-objectdescriptor-xml;windows_formatname="Star Object Descriptor (XML)"');
  }

finalization
//  CellClipboard.Free;
  if ComboColors <> nil then ComboColors.Free;


end.
