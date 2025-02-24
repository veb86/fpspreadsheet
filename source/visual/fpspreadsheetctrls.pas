{@@ ----------------------------------------------------------------------------
  Unit fpspreadsheetctrls implements some **visual controls** which help
  to create a spreadsheet application without writing too much code.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.

  EXAMPLE OF USAGE:
  * Add a _WorkbookSource_ component to the form.
  * Add a _WorksheetTabControl_
  * Add a _WorksheetGrid_ (from unit fpspreadsheetgrid)
  * Link their _WorkbookSource_ properties to the added
    _WorkbookSource_ component
  * Set the property _FileName_ of the _WorkbookSource_ to a
    spreadsheet file.

  As a result, the _WorksheetTabControl_ displays tabs for each worksheet
  in the file, and the _WorksheetGrid_ displays the worksheet according
  to the selected tab.
-------------------------------------------------------------------------------}
unit fpspreadsheetctrls;

{$MODE objfpc}{$H+}
{$include ..\fps.inc}

interface

uses
  LCLType, LCLIntf, LCLProc, LCLVersion, LMessages, LResources,
  Classes, Types, Graphics, SysUtils, Controls, StdCtrls, ComCtrls, ValEdit, ActnList,
  fpstypes, fpspreadsheet;

const
  {@@ User-defined message for input validation }
  UM_VALIDATEINPUT = LM_USER + 100;

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
    lniWorksheetRename, lniWorksheetZoom,
    lniCell, lniSelection, lniAbortSelection, lniRow, lniCol, lniChart);
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
    FOptions: TsWorkbookOptions;
    FUsesBuiltinWorkbook: Boolean;
    FOnError: TsWorkbookSourceErrorEvent;
    FOnQueryPassword: TsOnQueryPassword;

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
    procedure ColChangedHandler(Sender: TObject; ACol: Cardinal);
    procedure RowChangedHandler(Sender: TObject; ARow: Cardinal);
//    procedure WorkbookChangedPaletteHandler(Sender: TObject);
    procedure WorkbookOpenedHandler(Sender: TObject);
    procedure WorksheetAddedHandler(Sender: TObject; ASheet: TsWorksheet);
    procedure WorksheetChangedHandler(Sender: TObject; ASheet: TsWorksheet);
    procedure WorksheetRemovedHandler(Sender: TObject; ASheetIndex: Integer);
    procedure WorksheetRemovingHandler(Sender: TObject; AWorksheet: TsWorksheet);
    procedure WorksheetRenamedHandler(Sender: TObject; AWorksheet: TsWorksheet);
    procedure WorksheetSelectedHandler(Sender: TObject; AWorksheet: TsWorksheet);
    procedure WorksheetZoomHandler(Sender: TObject);

  protected
    procedure AbortSelection;
    function DoQueryPassword: String;
    procedure DoShowError(const AErrorMsg: String);
    procedure InternalCreateNewWorkbook(AWorkbook: TsWorkbook = nil);
    procedure InternalLoadFromFile(AFileName: string; AAutoDetect: Boolean;
      AFormatID: TsSpreadFormatID; AWorksheetIndex: Integer; APassword: String);
    procedure InternalLoadFromWorkbook(AWorkbook: TsWorkbook;
      AWorksheetIndex: Integer = -1);
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

    procedure LoadFromProtectedSpreadsheetFile(AFileName: String;
      AFormatID: TsSpreadFormatID; APassword: String; AWorksheetIndex: Integer = -1);
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = -1); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormatID: TsSpreadFormatID = sfidUnknown; AWorksheetIndex: Integer = -1); overload;
    procedure LoadFromWorkbook(AWorkbook: TsWorkbook; AWorksheetIndex: Integer = -1);

    procedure SaveToSpreadsheetFile(AFileName: string;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormatID: TsSpreadFormatID;
      AOverwriteExisting: Boolean = true); overload;

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

    property UsesBuiltinWorkbook: Boolean read FUsesBuiltinWorkbook;

  published
    {@@ Automatically detects the fileformat when loading the spreadsheet file
      specified by FileName }
    property AutoDetectFormat: Boolean read FAutoDetectFormat write FAutoDetectFormat default true;
    {@@ File format of the next spreadsheet file to be loaded by means of the
      Filename property. Not used when AutoDetectFormat is TRUE.
      Note that if FileFormat is sfUser then the format ID must be specified at
      runtime. }
    property FileFormat: TsSpreadsheetFormat read GetFileFormat write SetFileFormat default sfOOXML;
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
    {@@ Event fired when a password is required. Handler must return the pwd. }
    property OnQueryPassword: TsOnQueryPassword read FOnQueryPassword write FOnQueryPassword;
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
    FShowAllSheets: Boolean;
    procedure SetShowAllSheets(AValue: Boolean);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Change; override;
    procedure GetSheetList(AList: TStrings);
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    {@@ The worksheet names of this workbook are currently displayed as tabs of the TabControl. }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Identifies the worksheet which corresponds to the selected tab }
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property ShowAllSheets: boolean read FShowAllSheets write SetShowAllSheets default true;
    {@@ Link to the WorkbookSource which provides the data. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
  end;

  { TsWorksheetIndicator }

  {@@ TsWorksheetIndicator is a combobox control which displays the sheets of the
      workbook currently loaded by the WorkbookSource. }
  TsWorksheetIndicator = class(TCustomComboBox, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FLockCount: Integer;
    FShowAllSheets: Boolean;
    procedure SetShowAllSheets(AValue: Boolean);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Change; override;
    procedure GetSheetList(AList: TStrings);
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    property Workbook: TsWorkbook read GetWorkbook;
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    property ShowAllSheets: boolean read FShowAllSheets write SetShowAllSheets default true;
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


  { TsCellEdit }

  {@@ TsCellEdit allows to edit the content or formula of the active cell of a
    worksheet, simular to Excel's cell editor above the cell grid. }
  TsCellEdit = class(TMemo, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FShowHTMLText: Boolean;
    FOldText: String;
    FFormulaError: Boolean;
    FRefocusing: TObject;
    FRefocusingCol, FRefocusingRow: Cardinal;
    FRefocusingSelStart: Integer;
    function GetSelectedCell: PCell;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
    procedure ValidateInput(var Msg: TLMessage); message UM_VALIDATEINPUT;
    procedure WMKillFocus(var AMessage: TLMKillFocus); message LM_KILLFOCUS;
  protected
    FEditText: String;
    function CanEditCell(ACell: PCell): Boolean; overload;
    function CanEditCell(ARow, ACol: Cardinal): Boolean; overload;
    procedure DoEnter; override;
    procedure DoExit; override;
    function DoValidText(const AText: String): Boolean; virtual;
    function ValidFormula(AFormula: String; out AErrMsg: String): Boolean;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure KeyUp(var Key: Word; Shift: TShiftState); override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure ShowCell(ACell: PCell); virtual;
    procedure WriteToCell;
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
    property WantReturns default false;
  end;


  { TsCellIndicator }

  {@@ TsCellIndicator displays the address of the currently active cell of the
    worksheet and workbook. Editing the address allows to jump to the corresponding
    cell. }
  TsCellIndicator = class(TCustomEdit, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure UpdateDisplay;
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

    // Inherited from TCustomEdit, published in TEdit. Omit those which are not needed
    property Action;
    property Align;
    property Alignment default taCenter;   // centered text by default
    property Anchors;
    property AutoSize;
    property AutoSelect;
    property BidiMode;
    property BorderSpacing;
    property BorderStyle;
//    property CharCase;
    property Color;
    property Constraints;
    property DragCursor;
    property DragKind;
    property DragMode;
//    property EchoMode;
    property Enabled;
    property Font;
    property HideSelection;
    property MaxLength;
//    property NumbersOnly;
    property ParentBidiMode;
    property OnChange;
    property OnChangeBounds;
    property OnClick;
    property OnContextPopup;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEditingDone;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseEnter;
    property OnMouseLeave;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseWheel;
    property OnMouseWheelDown;
    property OnMouseWheelUp;
    property OnResize;
    property OnStartDrag;
    property OnUTF8KeyPress;
    property ParentColor;
    property ParentFont;
    property ParentShowHint;
//    property PasswordChar;
    property PopupMenu;
    property ReadOnly;
    property ShowHint;
    property TabStop;
    property TabOrder;
    property Text;
//    property TextHint;
    property Visible;
  end;


  { TsCellFormatItem, TsFormatTarget }

  TsCellFormatItem = (cfiFontName, cfiFontSize, cfiFontColor, cfiBackgroundColor,
    cfiBorderColor);

  TsFormatTarget = (ftCell, ftRow, ftCol, ftDefault);


  { TsCellCombobox }

  TsColorNameEvent = procedure (Sender: TObject; AColor: TColor;
    out AColorName: String) of object;

 {@@ TsCellCombobox is a multi-purpose combobox for selection of formatting
     items of a cell }
  TsCellCombobox = class(TCustomCombobox, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FFormatItem: TsCellFormatItem;
    FColorRectOffset: Integer;
    FColorRectWidth: Integer;
    FFormatTarget: TsFormatTarget;
    FOnAddColors: TNotifyEvent;
    FOnGetColorName: TsColorNameEvent;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetColorRectOffset(AValue: Integer);
    procedure SetColorRectWidth(AValue: Integer);
    procedure SetFormatItem(AValue: TsCellFormatItem);
    procedure SetFormatTarget(AValue: TsFormatTarget);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    procedure ApplyFormatToCell(ARow, ACol: Cardinal); virtual;
    procedure ApplyFormatToCol(ACol: Cardinal); virtual;
    procedure ApplyFormatToDefault; virtual;
    procedure ApplyFormatToRow(ARow: Cardinal); virtual;
    procedure ApplyFormat(ARow, ACol: cardinal);
    procedure Change; override;
    procedure DrawItem(AIndex: Integer; ARect: TRect;
      AState: TOwnerDrawState); override;
    procedure ExtractFromCell(ARow, ACol: Cardinal); virtual;
    procedure ExtractFromCol(ACol: Cardinal); virtual;
    procedure ExtractFromDefault; virtual;
    procedure ExtractFromRow(ARow: Cardinal); virtual;
    procedure ExtractFromSheet;
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
    procedure AddColor(AColor: TsColor; AColorName: String);
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    {@@ Refers to the underlying workbook }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Refers to the underlying worksheet containing the displayed cell }
    property Worksheet: TsWorksheet read GetWorksheet;
  published
    {@@ Identifies the cell format property to be used in the combobox }
    property CellFormatItem: TsCellFormatItem read FFormatItem write SetFormatItem default cfiFontName;
    {@@ Margin around the color box }
    property ColorRectOffset: Integer read FColorRectOffset write SetColorRectOffset default 2;
    {@@ Width of the color box shown for the color-related format items }
    property ColorRectWidth: Integer read FColorRectWidth write SetColorRectWidth default 10;
    {@@ Determine whether the selected color applies to a cell, row, column or default format }
    property FormatTarget: TsFormatTarget read FFormatTarget write SetFormatTarget default ftCell;
    {@@ Link to the WorkbookSource which provides the workbook and worksheet. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    {@@ Event which adds the colors to the combobox }
    property OnAddColors: TNotifyEvent read FOnAddColors write FOnAddColors;
    {@@ Event to get a decent name of the colors of the combo }
    property OnGetColorName: TsColorNameEvent read FOnGetColorName write FOnGetColorName;

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
  TsInspectorMode = (imWorkbook, imWorksheet, imCellValue, imCellProperties,
    imRow, imCol);

  {@@ Inspector expanded nodes }
  TsInspectorExpandedNode = (ienFormatSettings, ienConditionalFormats,
    ienPageLayout, ienFonts, ienFormats, ienEmbeddedObj, ienImages,
    ienCryptoInfo);
  TsInspectorExpandedNodes = set of TsInspectorExpandedNode;

  {@@ TsSpreadsheetInspector displays all properties of a workbook, worksheet,
    cell content and cell formatting in a way similar to the Object Inspector
    of Lazarus. }
  TsSpreadsheetInspector = class(TValueListEditor, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FMode: TsInspectorMode;
    FExpanded: TsInspectorExpandedNodes;
    FCurrRow, FCurrCol: Integer;
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
    procedure UpdateCol(ACol: Integer; AStrings: TStrings); virtual;
    procedure UpdateFormatProperties(AFormatIndex: integer; AStrings: TStrings); virtual;
    procedure UpdateRow(ARow: Integer; AStrings: TStrings); virtual;
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
      read FExpanded write SetExpanded
      default [ienFormatSettings, ienConditionalFormats, ienPageLayout,
               ienFonts, ienFormats, ienEmbeddedObj, ienImages, ienCryptoInfo];
    {@@ inherited from TValueListEditor. Turns of the fixed column by default}
    property FixedCols default 0;
    {@@ inherited from TStringGrid, but not published in TValueListEditor. }
    property ExtendedColSizing;
  end;

function SpreadsheetFormatInClipboard: Boolean;

{$IFDEF LCL_FULLVERSION_LT_180}
function ScalePPI(ALength: Integer): Integer;
{$IFEND}


implementation

uses
  Math, StrUtils, TypInfo, Dialogs, Forms, Clipbrd,
  fpsStrings, fpsCrypto, fpsReaderWriter, fpsUtils, fpsNumFormat, fpsImages,
  fpsHTMLUtils, fpsExprParser;

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
  Returns @TRUE if the clipboard contains a format good for pasting into a
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


{$IFDEF LCL_FullVersion_LT_180}
function ScalePPI(ALength: Integer): Integer;
begin
  Result := MulDiv(ALength, Screen.PixelsPerInch, 96);
end;
{$IFEND}

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

{@@
  Adds a "non-existing" cell to the list. Such a cell is nil in the worksheet.
  Here it has ContentType = cctEmpty and UsedFormattingFields = [], i.e. it is
  an empty cell without formatting. }
function TsCellList.AddEmptyCell(ARow, ACol: Cardinal): Integer;
var
  cell: PCell;
begin
  New(cell);
  InitCell(nil, ARow, ACol, cell^);
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

  @param  AOwner  Component which is responsibile for destroying the WorkbookSource.
-------------------------------------------------------------------------------}
constructor TsWorkbookSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FListeners := TFPList.Create;
  FFileFormatID := ord(sfOOXML);
  FAutoDetectFormat := True;
  FUsesBuiltinWorkbook := true;
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

  @param(AListener  Component to be added to the listener list notified of
                    changes.)
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
  Event handler for the OnChangeCol event of TsWorksheet which is fired whenver
  a column width or column format changes.

  @param   Sender   Pointer to the worksheet
  @param   ACol     Index (in sheet notation) of the column changed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.ColChangedHandler(Sender: TObject;
  ACol: Cardinal);
begin
  if FWorksheet <> nil then
    NotifyListeners([lniCol], {%H-}Pointer(PtrInt(ACol)));
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
var
  dummycell: TCell;
begin
  dummycell.Row := ARow;
  dummycell.Col := ACol;
//  Unused(ARow, ACol);
  NotifyListeners([lniSelection], @dummycell);

  if FPendingOperation <> coNone then
  begin
    ExecutePendingOperation;
    FPendingOperation := coNone;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Event handler for the OnChangeRow event of TsWorksheet which is fired whenver
  a row width or row format changes.

  Adds the index of the affected row to the Data field of the notification event.

  @param   Sender   Pointer to the worksheet
  @param   ARow     Index (in sheet notation) of the row changed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.RowChangedHandler(Sender: TObject;
  ARow: Cardinal);
begin
  if FWorksheet <> nil then
    NotifyListeners([lniRow], {%H-}Pointer(PtrInt(ARow)));
end;


{@@ ----------------------------------------------------------------------------
  Creates a new empty workbook and adds a single worksheet
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.CreateNewWorkbook;
begin
  FFileName := '';
  FFileFormatID := sfidUnknown;
  InternalCreateNewWorkbook;
  FWorksheet := FWorkbook.AddWorksheet(Format(rsDefaultSheetName,[1]));
  SelectWorksheet(FWorksheet);
end;

function TsWorkbookSource.DoQueryPassword: String;
var
  crs: TCursor;
begin
  crs := Screen.Cursor;
  Screen.Cursor := crDefault;
  try
    if Assigned(FOnQueryPassword) then
      Result := FOnQueryPassword()
    else
      Result := InputBox('Password required to open workbook', 'Password', '');
  finally
    Screen.Cursor := crs;
  end;
end;

{@@ ----------------------------------------------------------------------------
  An error has occured during loading of the workbook. Shows a message box by
  default. But a different behavior can be obtained by means of the OnError
  event.

  @param(AErrorMsg  Error message text created by the workbook reader and to be
                    displayed in a messagebox or by means of the OnError
                    handler.)
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.DoShowError(const AErrorMsg: String);
begin
  if Assigned(FOnError) then
    FOnError(self, AErrorMsg)
  else
    MessageDlg(AErrorMsg, mtError, [mbOK], 0);
end;

{@@ ----------------------------------------------------------------------------
  Executes a "pending operation"
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.ExecutePendingOperation;
var
  destSelection: TsCellRangeArray = nil;
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
    FWorkbook := TsWorkbook.Create
  else
    FWorkbook := AWorkbook;
  FWorkbook.OnOpenWorkbook := @WorkbookOpenedHandler;
  FWorkbook.OnAddWorksheet := @WorksheetAddedHandler;
  FWorkbook.OnChangeWorksheet := @WorksheetChangedHandler;
  FWorkbook.OnRemoveWorksheet := @WorksheetRemovedHandler;
  FWorkbook.OnRemovingWorksheet := @WorksheetRemovingHandler;
  FWorkbook.OnRenameWorksheet := @WorksheetRenamedHandler;
  FWorkbook.OnSelectWorksheet := @WorksheetSelectedHandler;
  // Pass options to workbook
  SetOptions(FOptions);
end;

{@@ ----------------------------------------------------------------------------
  Internal loader for the spreadsheet file. Is called with various combinations
  of arguments from several procedures.

  @param(AFilename        Name of the spreadsheet file to be loaded.)
  @param(AAutoDetect      Instructs the loader to automatically detect the
                          file format from the file extension or by temporarily
                          opening the file in all available formats. Note that
                          an exception is raised in the IDE when an incorrect
                          format is tested.)
  @param(AFormatID        Identifier of the spreadsheet file format assumed
                          for the loader.
                          Is ignored when AAutoDetect is @false.)
  @param(AWorksheetIndex  Index of the worksheet to be selected after loading.)
  @param(APassword        Password to open encrypted workbook. Note: this is
                          supported only by ods and xlsx readers.)
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.InternalLoadFromFile(AFileName: string;
  AAutoDetect: Boolean; AFormatID: TsSpreadFormatID; AWorksheetIndex: Integer;
  APassword: String);
var
  book: TsWorkbook;
begin
  book := TsWorkbook.Create;
  try
    book.Options := FOptions;
    book.OnQueryPassword := @DoQueryPassword;
    if AAutoDetect then
      book.ReadFromFile(AFileName, APassword)
    else
      book.ReadFromFile(AFileName, AFormatID, APassword);
    InternalLoadFromWorkbook(book, AWorksheetIndex);
  except
    // book is normally used as current workbook. But it must be destroyed
    // if the file cannot be read.
    book.Free;
    raise;
  end;
end;

procedure TsWorkbookSource.InternalLoadFromWorkbook(AWorkbook: TsWorkbook;
  AWorksheetIndex: Integer = -1);
begin
  AWorkbook.DisableNotifications;

  if AWorkbook <> FWorkbook then
    InternalCreateNewWorkbook(AWorkbook)
  else
    SetOptions(FOptions);
  WorkbookOpenedHandler(self);

  if AWorksheetIndex = -1 then
  begin
    if FWorkbook.ActiveWorksheet <> nil then
      AWorksheetIndex := FWorkbook.GetWorksheetIndex(FWorkbook.ActiveWorksheet) else
      AWorksheetIndex := 0;
  end;
  SelectWorksheet(FWorkbook.GetWorksheetByIndex(AWorksheetIndex));

  FUsesBuiltinWorkbook := false;

  AWorkbook.EnableNotifications;

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

  @param(AFilename        Name of the spreadsheet file to be loaded.)
  @param(AFormat          Spreadsheet file format assumed for the file.)
  @param(AWorksheetIndex  Index of the worksheet to be selected after loading.
                          (If empty then the active worksheet is loaded) )
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromSpreadsheetFile(AFileName: string;
  AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = -1);
begin
  if AFormat = sfUser then
    raise Exception.Create('[TsWorkbook.ReadFromFile] Don''t call this method for user-provided file formats.');
  LoadFromSpreadsheetFile(AFileName, ord(AFormat), AWorksheetIndex);
end;

{@@ ----------------------------------------------------------------------------
  Public loader of a spreadsheet file.

  Call this method for both built-in and user-provided file formats.

  If the workbook is password-protected the password is prompted by a dialog.

  @param(AFilename        Name of the spreadsheet file to be loaded.)
  @param(AFormatID        Identifier of the spreadsheet file format assumed
                          for the file.)
  @param(AWorksheetIndex  Index of the worksheet to be selected after loading.
                          (If empty then the active worksheet is loaded) )
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromSpreadsheetFile(AFileName: string;
  AFormatID: TsSpreadFormatID = sfidUnknown; AWorksheetIndex: Integer = -1);
var
  autodetect: Boolean;
begin
  autodetect := (AFormatID = sfidUnknown);
  InternalLoadFromFile(AFileName, autodetect, AFormatID, AWorksheetIndex, '');
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
  Uses an already existing workbook in the visual controls.

  IMPORTANT: THE CALLING ROUTINE MUST NOT DESTROY THE WORKBOOK, it is destroyed
  here by the TsWorkbookSource.
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromWorkbook(AWorkbook: TsWorkbook;
  AWorksheetIndex: Integer = -1);
begin
  InternalLoadFromWorkbook(AWorkbook, AWorksheetIndex);
end;

{@@ ----------------------------------------------------------------------------
  Public loader of a spreadsheet file.

  Should be called in case of password-protected files when the password is
  already known and no password dialog should appear.

  Call this method for both built-in and user-provided file formats.

  @param(AFilename        Name of the spreadsheet file to be loaded.)
  @param(AFormatID        Identifier of the spreadsheet file format assumed
                          for the file.)
  @param(AWorksheetIndex  Index of the worksheet to be selected after loading.
                          (If empty then the active worksheet is loaded) )
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.LoadFromProtectedSpreadsheetFile(AFileName: String;
  AFormatID: TsSpreadFormatID; APassword: String; AWorksheetIndex: Integer = -1);
begin
  InternalLoadFromFile(AFileName, false, AFormatID, AWorksheetIndex, APassword);
end;

{@@ ----------------------------------------------------------------------------
  Notifies listeners of workbook, worksheet, cell, or selection changes.
  The changed item is identified by the parameter AChangedItems.

  @param(AChangedItems  A set of elements lniWorkbook, lniWorksheet,
                        lniCell, lniSelection which indicate which item has
                        changed.)
  @param(AData          Additional information on the change. Is used only for
                        lniCell and points to the cell having a changed value
                        or formatting.)
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
    if (C <> nil) then
    begin
      if C.GetInterface(GUID_SpreadsheetControl, I) then
        I.ListenerNotification(AChangedItems, AData)
      else
        raise Exception.CreateFmt('[TsWorkbookSource.NotifyListeners] Class %s is not prepared to be a spreadsheet listener.',
          [C.ClassName]);
    end;
  end;

  // Cleanup listener list from removed listeners (= set to nil) while
  // NotifyListeners was running
  for j := FListeners.Count-1 downto 0 do
    if FListeners[j] = nil then
      FListeners.Delete(j);
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
      FListeners[j] := nil;
      // Do not delete the listener here (FListeners.Delete(j)) because
      // RemoveListeners may be called while NotifyListeners is still running.
      // The problem can be that a chart may destroy a listening chart source
      // which would trigger RemoveListener. If the chart source then would be
      // deleted from the list the NotifiyListeners loop would access
      // unallocated memory --> crash
      if C <> nil then
      begin
        if C.GetInterface(GUID_SpreadsheetControl, I) then
          I.RemoveWorkbookSource
        else
          raise Exception.CreateFmt('Class %s not prepared for listening.',[AListener.ClassName]);
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the workbook of the WorkbookSource component to a spreadsheet file.

  Call this method only for built-in file formats.

  @param(AFileName          Name of the file to which the workbook is to be
                            saved.)
  @param(AFormat            Spreadsheet file format in which the file is to be
                            saved.)
  @param(AOverwriteExisting If the file already exists, it is overwritten in
                            the case of AOverwriteExisting = @true, or an
                            exception is raised otherwise.)
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

  @param(AFileName          Name of the file to which the workbook is to be
                            saved.)
  @param(AFormatID          Identifier of the spreadsheet file format in which
                            the file is to be saved.)
  @param(AOverwriteExisting If the file already exists, it is overwritten in
                            the case of AOverwriteExisting = @true, or an
                            exception is raised otherwise.)
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

  @param(AFileName          Name of the file to which the workbook is to be
                            saved.
                            If the file format is not known is is written
                            as BIFF8/XLS. )
  @param(AOverwriteExisting If this file already exists it is overwritten if
                            AOverwriteExisting = @true, or an exception is
                            raised if AOverwriteExisting = @false.)
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
var
  AutoCalcChanged: Boolean;
begin
  AutoCalcChanged := (FOptions * [boAutoCalc]) <> (AValue* [boAutoCalc]);
  FOptions := AValue;
  if Workbook <> nil then begin
    Workbook.Options := FOptions;
    if AutoCalcChanged and (boAutoCalc in FOptions) then
      Workbook.CalcFormulas;
  end;
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
  s: String;

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
      if FWorksheet.GetSelectionCount = 1 then begin
        s := FWorksheet.ReadAsText(FWorksheet.ActiveCellRow, FWorksheet.ActiveCellCol);
        if s <> '' then begin
          stream.Write(s[1], Length(s));
          stream.WriteByte(0);
          stream.Position := 0;
          Clipboard.AddFormat(CF_TEXT, stream);
          (stream as TMemoryStream).Clear;
        end;
      end else
      begin
        CsvParams.Delimiter := #9;
        CopyToClipboard(stream, sfCSV, CF_TEXT);
        CSVParams := savedCSVParams;
      end;

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

  If ATranspose is @TRUE then rows and columns are interchanged.
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
    SelectWorksheet(FWorkbook.GetFirstWorksheet)
  else
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

  @param(Sender       Points to the workbook from which the sheet has been
                      deleted.)
  @param(ASheetIndex  Index of the sheet that was deleted. The sheet itself
                      does not exist any more.)
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
  if ASheetIndex > -1 then
    sheet := FWorksheet
  else
    sheet := nil;  // all sheets were removed

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
    FWorksheet.OnChangeCol := @ColChangedHandler;
    FWorksheet.OnChangeFont := @CellFontChangedHandler;
    FWorksheet.OnChangeRow := @RowChangedHandler;
    FWorksheet.OnSelectCell := @CellSelectedHandler;
    FWorksheet.OnZoom := @WorksheetZoomHandler;
    NotifyListeners([lniWorksheet]);
    FWorksheet := AWorksheet;  // !!!!!
    if FWorksheet.ActiveCellRow = UNASSIGNED_ROW_COL_INDEX then
      r := FWorksheet.TopPaneHeight else
      r := FWorksheet.ActiveCellRow;
    if FWorksheet.ActiveCellCol = UNASSIGNED_ROW_COL_INDEX then
      c := FWorksheet.LeftPaneWidth else
      c := FWorksheet.ActiveCellCol;
    SelectCell(r, c);
  end else
    NotifyListeners([lniWorksheet]);
end;

{@@ ----------------------------------------------------------------------------
  Event handler called whenever the workbook is zoomed
-------------------------------------------------------------------------------}
procedure TsWorkbookSource.WorksheetZoomHandler(Sender: TObject);
begin
  NotifyListeners([lniWorksheetZoom], FWorksheet);
end;


{------------------------------------------------------------------------------}
{                            TsWorkbookTabControl                              }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the WorkbookTabControl.
-------------------------------------------------------------------------------}
constructor TsWorkbookTabControl.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FShowAllSheets := true;
end;

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
var
  sheetName: String;
begin
  if (FWorkbookSource <> nil) and (FLockCount = 0) then begin
    sheetName := Tabs[TabIndex];
    FWorkbookSource.SelectWorksheet(Workbook.GetWorksheetByName(sheetName));
  end;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Populates a (string) list with the names of the workbook's sheet names.
  If ShowHiddenSheets is false hidden worksheets are skipped.
  Is called whenever the workbook changes.
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.GetSheetList(AList: TStrings);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  AList.Clear;
  if Workbook <> nil then
    for i:=0 to Workbook.GetWorksheetCount-1 do
    begin
      sheet := Workbook.GetWorksheetByIndex(i);
      if FShowAllSheets or not (soHidden in sheet.Options) then
        AList.Add(sheet.Name);
    end;
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

  @param(AChangedItems  Set with elements identifying whether workbook,
                        worksheet cell content or cell formatting has changed.)
  @param(AData          Additional data, not used here.)
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  Unused(AData);

  // Workbook changed: new workbook, worksheet added/renamed/deleted
  if (AChangedItems * [lniWorkbook, lniWorksheetAdd, lniWorksheetRemove, lniWorksheetRename] <> []) then
  begin
    inc(FLockCount);    // avoid WorkbookSelect message when adding each tab
    GetSheetList(Tabs);
    if (lniWorkbook in AChangedItems) and (Workbook <> nil) then
    begin
      i := Tabs.IndexOf(Workbook.ActiveWorksheet.Name);
      if i > -1 then TabIndex := i else TabIndex := 0
    end else
    if (lniWorksheetAdd in AChangedItems) then
      TabIndex := Tabs.Count-1
    else
    if (lniWorksheetRename in AChangedItems) then begin
      sheet := TsWorksheet(AData);
      TabIndex := Tabs.IndexOf(sheet.Name);
    end;
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
  Setter method for the property ShowHiddenSheets
-------------------------------------------------------------------------------}
procedure TsWorkbookTabControl.SetShowAllSheets(AValue: Boolean);
var
  idx, i: Integer;
  sheet: TsWorksheet;
begin
  if AValue = FShowAllSheets then
    exit;
  FShowAllSheets := AValue;
  idx := -1;
  // Find tabindex of next visible sheet
  if not FShowAllSheets and (Workbook <> nil) then begin
    for i:=0 to Workbook.GetWorksheetCount-1 do begin
      sheet := Workbook.GetWorksheetByIndex(i);
      if sheet = Worksheet then
        break;
      if not (soHidden in sheet.Options) then inc(idx);
    end;
    i := idx;
    while (sheet <> nil) and (soHidden in sheet.Options) do begin
      inc(i);
      sheet := Workbook.GetWorksheetByIndex(i);
    end;
    if sheet = nil then begin
      i := idx;
      while (sheet <> nil) and (soHidden in sheet.Options) do begin
        dec(i);
        sheet := Workbook.GetWorksheetByIndex(i);
      end;
      if sheet = nil then idx := -1;
    end;
  end;
  Change;
  if (not FShowAllSheets) then
    TabIndex := idx;
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
{                            TsWorksheetIndicator                              }
{------------------------------------------------------------------------------}

constructor TsWorksheetIndicator.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FShowAllSheets := true;
end;

destructor TsWorksheetIndicator.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

procedure TsWorksheetIndicator.Change;
var
  sheetName: String;
begin
  if (FWorkbookSource <> nil) and (FLockCount = 0) then begin
    sheetName := Items[ItemIndex];
    FWorkbookSource.SelectWorksheet(Workbook.GetWorksheetByName(sheetName));
  end;
  inherited;
end;

procedure TsWorksheetIndicator.GetSheetList(AList: TStrings);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  AList.Clear;
  if Workbook <> nil then
    for i:=0 to Workbook.GetWorksheetCount-1 do
    begin
      sheet := Workbook.GetWorksheetByIndex(i);
      if FShowAllSheets or not (soHidden in sheet.Options) then
        AList.Add(sheet.Name);
    end;
end;

function TsWorksheetIndicator.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

function TsWorksheetIndicator.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

procedure TsWorksheetIndicator.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  Unused(AData);

  // Workbook changed: new workbook, worksheet added/renamed/deleted
  if (AChangedItems * [lniWorkbook, lniWorksheetAdd, lniWorksheetRemove, lniWorksheetRename] <> []) then
  begin
    inc(FLockCount);    // avoid WorkbookSelect message when adding each tab
    GetSheetList(Items);
    if (lniWorkbook in AChangedItems) and (Workbook <> nil) then
    begin
      i := Items.IndexOf(Workbook.ActiveWorksheet.Name);
      if i > -1 then ItemIndex := i else ItemIndex := 0
    end else
    if (lniWorksheetAdd in AChangedItems) then
      ItemIndex := Items.Count-1
    else
    if (lniWorksheetRename in AChangedItems) then begin
      sheet := TsWorksheet(AData);
      ItemIndex := Items.IndexOf(sheet.Name);
    end;
    dec(FLockCount);
  end;

  // Worksheet selected
  if (lniWorksheet in AChangedItems) and (Worksheet <> nil) then
  begin
    i := Items.IndexOf(Worksheet.Name);
    if i <> ItemIndex then
      ItemIndex := i;
  end;
end;

procedure TsWorksheetIndicator.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

procedure TsWorksheetIndicator.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

procedure TsWorksheetIndicator.SetShowAllSheets(AValue: Boolean);
var
  idx, i: Integer;
  sheet: TsWorksheet;
begin
  if AValue = FShowAllSheets then
    exit;
  FShowAllSheets := AValue;
  idx := -1;
  // Find ItemIndex of next visible sheet
  if not FShowAllSheets and (Workbook <> nil) then begin
    for i:=0 to Workbook.GetWorksheetCount-1 do begin
      sheet := Workbook.GetWorksheetByIndex(i);
      if sheet = Worksheet then
        break;
      if not (soHidden in sheet.Options) then inc(idx);
    end;
    i := idx;
    while (sheet <> nil) and (soHidden in sheet.Options) do begin
      inc(i);
      sheet := Workbook.GetWorksheetByIndex(i);
    end;
    if sheet = nil then begin
      i := idx;
      while (sheet <> nil) and (soHidden in sheet.Options) do begin
        dec(i);
        sheet := Workbook.GetWorksheetByIndex(i);
      end;
      if sheet = nil then idx := -1;
    end;
  end;
  Change;
  if (not FShowAllSheets) then
    ItemIndex := idx;
end;

procedure TsWorksheetIndicator.SetWorkbookSource(AValue: TsWorkbookSource);
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

function TsCellEdit.CanEditCell(ACell: PCell): Boolean;
begin
  if Worksheet.IsMerged(ACell) then
    ACell := Worksheet.FindMergeBase(ACell);
  Result := not (
    Worksheet.IsProtected and
    (spCells in Worksheet.Protection) and
    ((ACell = nil) or (cpLockcell in Worksheet.ReadCellProtection(ACell)))
  );
end;

function TsCellEdit.CanEditCell(ARow, ACol: Cardinal): Boolean;
var
  cell: PCell;
begin
  cell := Worksheet.FindCell(ARow, ACol);
  Result := CanEditCell(cell);
end;

procedure TsCellEdit.DoEnter;
begin
  if Worksheet = nil then
    exit;

  if not CanEditCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol) then
  begin
    MessageDlg('This cell is protected from editing. Unlock worksheet protection '+
      'before continuing.', mtInformation, [mbOK], 0);
    Abort;
  end;

  if FRefocusing = self then
    FRefocusing := nil;

  inherited;
end;

{@@ ----------------------------------------------------------------------------
  DoExit is called when the edit has lost focus. Posts a message to the end of
  the message queue in order to complete the focus change operation and to
  trigger validation of the edit afterwards (which will restore the edit as
  focused control if validation fails.

  Source of the idea:
  https://community.embarcadero.com/article/technical-articles/149-tools/12766-validating-input-in-tedit-components
-------------------------------------------------------------------------------}
procedure TsCellEdit.DoExit;
begin
  if FRefocusing = nil then begin
    // Remember current text in editor...
    FEditText := Text;
    // ... as well as currently selected cell.
    FRefocusingRow := Worksheet.ActiveCellRow;
    FRefocusingCol := Worksheet.ActiveCellCol;
    FRefocusingSelStart := SelStart;
    // Initiate validation of current input
    PostMessage(Handle, UM_VALIDATEINPUT, 0, LParam(Self));
  end;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Validation function for the current text in the edit control called by the
  handler of the message posted in DoExit. Checks whether a formula is valid.
  If not, the currently selected cell and the edit's Text are restored, and an
  error message is displayed.
-------------------------------------------------------------------------------}
function TsCellEdit.DoValidText(const AText: String): Boolean;
var
  err: String;
begin
  Result := ValidFormula(AText, err);
  Worksheet.SelectCell(FRefocusingRow, FRefocusingCol);
  Text := AText;  // restore orig text lost by interaction with grid
  SelStart := FRefocusingSelStart;
  if Result then
    WriteToCell
  else
    MessageDlg(err, mtError, [mbOK], 0);
end;

{@@ ----------------------------------------------------------------------------
  EditingDone is called when the user presses the RETURN key to finish editing.
  If the current Text is an invalid formula an error message is displayed and
  nothing else happens. Otherwise, however, the edited text is written to the
  worksheet which tries to figure out the data type.
-------------------------------------------------------------------------------}
procedure TsCellEdit.EditingDone;
var
  s: String;
begin
  if (Worksheet = nil) then
    exit;

  if not ValidFormula(Text, s) then begin
    MessageDlg(s, mtError, [mbOK], 0);
    exit;
  end;

  WriteToCell;

  inherited;
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
  Inherited KeyDown method. Overridden here in order to be able to restore the
  old cell content if ESC is pressed.
-------------------------------------------------------------------------------}
procedure TsCellEdit.KeyDown(var Key: Word; Shift : TShiftState);
var
  selpos: Integer;
  errMsg: String;
begin
  FFormulaError := false;
  case Key of
    VK_RETURN:
      if not ValidFormula(Text, errMsg) then begin
        Key := 0;
        FFormulaError := true;
        MessageDlg(errMsg, mtError, [mbOK], 0);
      end;
    VK_ESCAPE:
      begin
        Key := 0;
        selpos := SelStart;
        Lines.Text := FOldText;
        SelStart := selpos;
        exit;
      end;
  end;
  inherited;
end;

procedure TsCellEdit.KeyUp(var Key: Word; Shift : TShiftState);
begin
  if FFormulaError and (Key = VK_RETURN) then
    Key := 0;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which item
  of the spreadsheet has changed.
  Responds to selection and cell changes by updating the cell content.

  @param(AChangedItems  Set with elements identifying whether workbook, worksheet
                        cell content or cell formatting has changed.)
  @param(AData          If AChangedItems contains nliCell then AData points to
                        the modified cell.)
-------------------------------------------------------------------------------}
procedure TsCellEdit.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  cell: PCell;
begin
  if (FWorkbookSource = nil) or (FRefocusing = self) then
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
  hideformula: Boolean;
begin
  if (FWorkbookSource <> nil) and (ACell <> nil) then
  begin
    hideformula := Worksheet.IsProtected and (spCells in Worksheet.Protection) and
      (cpHideFormulas in Worksheet.ReadCellProtection(ACell));
    s := Worksheet.ReadFormulaAsString(ACell, true);
    if (s <> '') then begin
      if hideformula then
        s := '(Formula hidden)'
      else
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
          Lines.Text := Worksheet.ReadAsText(ACell);
      end;
  end else
    Clear;

  FOldText := Lines.Text;

  ReadOnly := not CanEditCell(ACell);
end;

{@@ ----------------------------------------------------------------------------
  Validation routine of current input initiated when the edit lost focus.
  If input is not correct (DoValidText) then state before focus change is
  re-established.
-------------------------------------------------------------------------------}
procedure TsCellEdit.ValidateInput(var Msg: TLMessage);
var
  s: String;
begin
  if TControl(Msg.lParam) is TsCellEdit then begin
    s := TsCellEdit(Msg.lParam).FEditText;
    if not DoValidText(s) then begin
      FRefocusing := TControl(Msg.lParam);          // Avoid an endless loop
      TWinControl(Msg.lParam).SetFocus;             // Set focus back
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks valididty of the provided formula and creates a corresponding
  error message.

  It is assumed that the formula is localized.

  @Returns(@TRUE if the provided string is a valid formula or no formula, @FALSE
  otherwise. In the latter case an error message string is returned as well.)
-------------------------------------------------------------------------------}
function TsCellEdit.ValidFormula(AFormula: String; out AErrMsg: String): Boolean;
var
  parser: TsSpreadsheetParser;
  cell: PCell;
begin
  Result := true;
  AErrMsg := '';

  if Assigned(Worksheet) and (AFormula <> '') and (AFormula[1] = '=') then
  begin
    cell := GetSelectedCell;
    parser := TsSpreadsheetParser.Create(Worksheet, cell^.Row, cell^.Col);
    try
      try
        parser.Expression[fdLocalized] := AFormula;
      except
        on E: Exception do begin
          AErrMsg := E.Message;
          Result := false;
        end;
      end;
    finally
      parser.Free;
    end;
  end;
end;

procedure TsCellEdit.WMKillFocus(var AMessage: TLMKillFocus);
begin
  Unused(AMessage);
  // Override inherited behavior because we don't want to call EditingDone
  // here.
end;

{@@ ----------------------------------------------------------------------------
  Writes the current edit text to the cell

  @Note  All validation checks already have been performed.
-------------------------------------------------------------------------------}
procedure TsCellEdit.WriteToCell;
var
  cell: PCell;
  s: String;
begin
  cell := Worksheet.GetCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  if Worksheet.IsMerged(cell) then
    cell := Worksheet.FindMergeBase(cell);
  s := Lines.Text;
  if (s <> '') and (s[1] = '=') then
    Worksheet.WriteFormula(cell, Copy(s, 2, Length(s)), true)
  else
    Worksheet.WriteCellValueAsString(cell, s);
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

  @param(AChangedItems  Set with elements identifying whether workbook, worksheet
                        cell or selection has changed. Only the latter element
                        is considered by the cell indicator.)
  @param(AData          If AChangedItems contains nliCell then AData points to
                        the modified cell.)
-------------------------------------------------------------------------------}
procedure TsCellIndicator.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
begin
  Unused(AData);
  if (lniSelection in AChangedItems) and (Worksheet <> nil) then
    UpdateDisplay;
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

procedure TsCellIndicator.UpdateDisplay;
var
  sel: TsCellRangeArray;
  s: String;
  rng: TsCellRange;
  numrows, numcols: Integer;
  r, c: Cardinal;
  cell: PCell;
begin
  r := Worksheet.ActiveCellRow;
  c := Worksheet.ActiveCellCol;
  cell := Worksheet.FindCell(r, c);
  if cell <> nil then begin
    cell := Worksheet.FindMergeBase(cell);
    if cell <> nil then begin
      r := cell^.Row;
      c := cell^.Col;
    end;
  end;
  s := GetCellString(r, c);
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
  Adds a named color to the combobox items
-------------------------------------------------------------------------------}
procedure TsCellCombobox.AddColor(AColor: TsColor; AColorName: String);
var
  noText: Boolean;
begin
  if (FFormatItem in [cfiFontColor, cfiBackgroundColor, cfiBorderColor]) then
  begin
    noText := (FColorRectWidth = -1);
    Items.AddObject(StrUtils.IfThen(noText, '', AColorName), TObject(PtrInt(AColor)));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Apply the selected format style to the cell, column, row or default format
  (depending in FFormatTarget)
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ApplyFormat(ARow, ACol: Cardinal);
begin
  case FFormatTarget of
    ftCell    : ApplyFormatToCell(ARow, ACol);
    ftCol     : ApplyFormatToCol(ACol);
    ftRow     : ApplyFormatToRow(ARow);
    ftDefault : ApplyformatToDefault;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Applies the format to a cell. Override according to the format item for
  which the combobox is responsible.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ApplyFormatToCell(ARow, ACol: Cardinal);
var
  fnt: TsFont;
  clr: TColor;
  cell: PCell;
begin
  if (Worksheet = nil) then
    exit;

  // Find cell at this location. Create a new cell here, if required.
  cell := Worksheet.GetCell(ARow, ACol);

  if Worksheet.IsMerged(cell) then
    cell := Worksheet.FindMergeBase(cell);

  case FFormatItem of
    cfiFontName:
      if Text <> '' then
      begin
        fnt := Worksheet.ReadCellFont(cell);
        Worksheet.WriteFont(cell, Text, fnt.Size, fnt.Style, fnt.Color);
      end;
    cfiFontSize:
      if Text <> '' then
      begin
        fnt := Worksheet.ReadCellFont(cell);
        Worksheet.WriteFont(cell, fnt.FontName, StrToFloat(Text), fnt.Style, fnt.Color);
      end;
    cfiFontColor:
      if ItemIndex > -1 then
      begin
        fnt := Worksheet.ReadCellFont(cell);
        clr := PtrInt(Items.Objects[ItemIndex]);
        Worksheet.WriteFont(cell, fnt.FontName, fnt.Size, fnt.style, clr);
      end;
    cfiBackgroundColor:
      if ItemIndex <= 0 then
        Worksheet.WriteBackgroundColor(cell, scTransparent)
      else
      begin
        clr := PtrInt(Items.Objects[ItemIndex]);
        Worksheet.WriteBackgroundColor(cell, clr);
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatCombobox.ApplyFormatToCell] Unknown format item');
  end;
end;

{@@ ----------------------------------------------------------------------------
  Applies the format to a column format record.
  Override according to the format item for which the combobox is responsible.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ApplyFormatToCol(ACol: Cardinal);
var
  fnt: TsFont;
  col: PCol;
  fmt: PsCellFormat;
  idx: Integer;
  clr: TsColor;
begin
  if (Worksheet = nil) then
    exit;

  // Find column record having the specified index. Create new record if required.
  col := Worksheet.GetCol(ACol);
  fmt := Workbook.GetPointerToCellFormat(col^.FormatIndex);

  case FFormatItem of
    cfiFontName:
      if Text <> '' then
      begin
        fnt := Workbook.GetFont(fmt^.FontIndex);
        fnt.FontName := Text;
        fmt^.FontIndex := Workbook.AddFont(fnt);
        Worksheet.WriteColFormatIndex(ACol, Workbook.AddCellFormat(fmt^));
      end;
    cfiFontSize:
      if Text <> '' then
      begin
        fnt := Workbook.GetFont(fmt^.FontIndex);
        fnt.Size := StrToFloat(Text);
        fmt^.FontIndex := Workbook.AddFont(fnt);
        Worksheet.WriteColFormatIndex(ACol, Workbook.AddCellFormat(fmt^));
      end;
    cfiFontColor:
      if ItemIndex > -1 then
      begin
        fnt := Workbook.GetFont(fmt^.FontIndex);
        fnt.Color := PtrInt(Items.Objects[ItemIndex]);
        fmt^.FontIndex := Workbook.AddFont(fnt);
        Worksheet.WriteColFormatIndex(ACol, Workbook.AddCelLFormat(fmt^));
      end;
    cfiBackgroundColor:
      begin
        if ItemIndex <= 0 then
          idx := Worksheet.ChangeBackground(col^.FormatIndex, fsNoFill, scTransparent, scTransparent)
        else
        begin
          clr := PtrInt(Items.Objects[ItemIndex]);
          idx := Worksheet.ChangeBackground(col^.FormatIndex, fsSolidFill, clr, clr);
        end;
        Worksheet.WriteColFormatIndex(ACol, idx);
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatCombobox.ApplyFormatToCol] Unknown format item');
  end;
end;

procedure TsCellCombobox.ApplyFormatToDefault;
var
  fnt: TsFont;
  fmt: PsCellFormat;
begin
  fmt := Workbook.GetPointerToCellFormat(0);
  case FFormatItem of
    cfiFontName:
      if Text <> '' then begin
        fnt := Workbook.GetDefaultFont;
        Workbook.SetDefaultFont(Text, fnt.Size);
      end;
    cfiFontSize:
      if Text <> '' then begin
        fnt := Workbook.GetDefaultFont;
        Workbook.SetDefaultFont(fnt.FontName, StrToFloat(Text));
      end;
    cfiFontColor:
      ;  // No change of default font color
    cfiBackgroundColor:
      begin
        if ItemIndex <= 0 then
          fmt^.UsedFormattingFields := fmt^.UsedFormattingFields - [uffBackground]
        else
          fmt^.UsedFormattingfields := fmt^.UsedFormattingFields + [uffBackground];
        fmt^.Background.Style := fsSolidFill;
        fmt^.Background.BgColor := PtrInt(Items.Objects[ItemIndex]);;
        fmt^.Background.FgColor := fmt^.Background.BgColor;
      end;
    cfiBorderColor:
      ;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Applies the format to a row format record.
  Override according to the format item for which the combobox is responsible.
-------------------------------------------------------------------------------}
procedure TsCellCombobox.ApplyFormatToRow(ARow: Cardinal);
var
  fnt: TsFont;
  row: PRow;
  fmt: PsCellFormat;
  idx: Integer;
  clr: TsColor;
begin
  if (Worksheet = nil) then
    exit;

  // Find row record having the specified index. Create new record if required.
  row := Worksheet.GetRow(ARow);
  fmt := Workbook.GetPointerToCellFormat(row^.FormatIndex);

  case FFormatItem of
    cfiFontName:
      if Text <> '' then
      begin
        fnt := Workbook.GetFont(fmt^.FontIndex);
        fnt.FontName := Text;
        fmt^.FontIndex := Workbook.AddFont(fnt);
        Worksheet.WriteRowFormatIndex(ARow, Workbook.AddCellFormat(fmt^));
      end;
    cfiFontSize:
      if Text <> '' then
      begin
        fnt := Workbook.GetFont(fmt^.FontIndex);
        fnt.Size := StrToFloat(Text);
        fmt^.FontIndex := Workbook.AddFont(fnt);
        Worksheet.WriteRowFormatIndex(ARow, Workbook.AddCellFormat(fmt^));
      end;
    cfiFontColor:
      if ItemIndex > -1 then
      begin
        fnt := Workbook.GetFont(fmt^.FontIndex);
        fnt.Color := PtrInt(Items.Objects[ItemIndex]);
        fmt^.FontIndex := Workbook.AddFont(fnt);
        Worksheet.WriteRowFormatIndex(ARow, Workbook.AddCelLFormat(fmt^));
      end;
    cfiBackgroundColor:
      begin
        if ItemIndex <= 0 then
          idx := Worksheet.ChangeBackground(row^.FormatIndex, fsNoFill, scTransparent, scTransparent)
        else
        begin
          clr := PtrInt(Items.Objects[ItemIndex]);
          idx := Worksheet.ChangeBackground(row^.FormatIndex, fsSolidFill, clr, clr);
        end;
        Worksheet.WriteRowFormatIndex(ARow, idx);
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatCombobox.ApplyFormatToRow] Unknown format item');
  end;
end;

{@@ ----------------------------------------------------------------------------
  The text of the currently selected combobox item has been changed.
  Calls "ProcessValue" to changes the selected cells according to the
  Mode property by calling ApplyFormat.
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

     // Note: odPainted was renamed in LCL (v1.9) to odBackgroundPainted
     {$IF lcl_fullversion = 01090000}
      Exclude(AState, TOwnerDrawStateType(13));
       // ord(odPainted) = ord(odBackgroundPainted) = 13
     {$ELSEIF lcl_fullversion < 01090000}
      Exclude(AState, odPainted);
     {$ELSE}
      Exclude(AState, odBackgroundPainted);
     {$ENDIF}

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
procedure TsCellCombobox.ExtractFromCell(ARow, ACol: Cardinal);
var
  fnt: TsFont;
  clr: TsColor;
  cell: PCell;
begin
  cell := Worksheet.FindCell(ARow, ACol);
  if Worksheet.IsMerged(cell) then
    cell := Worksheet.FindMergeBase(cell);

  case FFormatItem of
    cfiFontName:
      begin
        fnt := Worksheet.ReadCellFont(cell);
        // No check for nil required because fnt is at least DefaultFont
        ItemIndex := Items.IndexOf(fnt.FontName);
        if ItemIndex = -1 then
          Text := fnt.FontName;
      end;
    cfiFontSize:
      begin
        fnt := Worksheet.ReadCellFont(cell);
        ItemIndex := Items.IndexOf(Format('%.0f', [fnt.Size]));
        if ItemIndex = -1 then
          Text := Format('%.2g', [fnt.Size]);
      end;
    cfiFontColor:
      begin
        fnt := Worksheet.ReadCellFont(cell);
        ItemIndex := Items.IndexOfObject(TObject(PtrInt(fnt.Color)));
        if ItemIndex = -1 then
          Text := GetColorName(fnt.Color);
      end;
    cfiBackgroundColor:
      begin
        clr := Worksheet.ReadBackgroundColor(cell);
        ItemIndex := Max(0, Items.IndexOfObject(TObject(PtrInt(clr))));
        if ItemIndex = -1 then
          Text := GetColorName(clr);
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatItem.ExtractFromCell] Unknown format item');
  end;
end;

procedure TsCellCombobox.ExtractFromCol(ACol: Cardinal);
var
  col: PCol;
  clr: TsColor;
  fnt: TsFont;
begin
  col := Worksheet.FindCol(ACol);
  case FFormatItem of
    cfiFontName:
      begin
        fnt := Worksheet.ReadColFont(col);
        // No check for nil required because fnt is at least DefaultFont
        ItemIndex := Items.IndexOf(fnt.FontName);
      end;
    cfiFontSize:
      begin
        fnt := Worksheet.ReadColFont(col);
        ItemIndex := Items.IndexOf(Format('%.0f', [fnt.Size]));
      end;
    cfiFontColor:
      begin
        fnt := Worksheet.ReadColFont(col);
        itemIndex := Items.IndexOfObject(TObject(PtrInt(fnt.Color)));
      end;
    cfiBackgroundColor:
      begin
        if col <> nil then clr := Worksheet.ReadBackgroundColor(col^.FormatIndex)
          else clr := Worksheet.ReadBackgroundColor(0);
        ItemIndex := Max(0, Items.IndexOfObject(TObject(PtrInt(clr))));
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatItem.ExtractFromCol] Unknown format item');
  end;
end;

procedure TsCellCombobox.ExtractFromDefault;
var
  fnt: TsFont;
  fmt: PsCellFormat;
begin
  fnt := Workbook.GetDefaultFont;
  case FFormatItem of
    cfiFontName:
      ItemIndex := Items.IndexOf(fnt.FontName);
    cfiFontSize:
      ItemIndex := Items.IndexOf(Format('%.0f', [fnt.Size]));
    cfiFontColor:
      ItemIndex := Items.IndexOfObject(TObject(PtrInt(fnt.Color)));
    cfiBackgroundColor:
      begin
        fmt := Workbook.GetPointerToCellFormat(0);
        if (uffBackground in fmt^.UsedFormattingFields) then
          ItemIndex := Items.IndexOfObject(TObject(PtrInt(fmt^.Background.BgColor)))
        else
          ItemIndex := Items.IndexOfObject(TObject(PtrInt(scTransparent)));
      end;
    cfiBorderColor:
      ;
  end;
end;

procedure TsCellCombobox.ExtractFromRow(ARow: Cardinal);
var
  row: PRow;
  clr: TsColor;
  fnt: TsFont;
begin
  row := Worksheet.FindRow(ARow);
  case FFormatItem of
    cfiFontName:
      begin
        fnt := Worksheet.ReadRowFont(row);
        // No check for nil required because fnt is at least DefaultFont
        ItemIndex := Items.IndexOf(fnt.FontName);
      end;
    cfiFontSize:
      begin
        fnt := Worksheet.ReadRowFont(row);
        ItemIndex := Items.IndexOf(Format('%.0f', [fnt.Size]));
      end;
    cfiFontColor:
      begin
        fnt := Worksheet.ReadRowFont(row);
        itemIndex := Items.IndexOfObject(TObject(PtrInt(fnt.Color)));
      end;
    cfiBackgroundColor:
      begin
        if row <> nil then clr := Worksheet.ReadBackgroundColor(row^.FormatIndex)
          else clr := Worksheet.ReadBackgroundColor(0);
        ItemIndex := Max(0, Items.IndexOfObject(TObject(PtrInt(clr))));
      end;
    cfiBorderColor:
      ;
    else
      raise Exception.Create('[TsCellFormatItem.ExtractFromCol] Unknown format item');
  end;
end;

procedure TsCellCombobox.ExtractFromSheet;
begin
  if (WorkbookSource = nil) or (Worksheet = nil) then
    exit;

  case FFormatTarget of
    ftCell:
      ExtractFromCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
    ftRow:
      ExtractFromRow(Worksheet.ActiveCellRow);
    ftCol:
      ExtractFromCol(Worksheet.ActiveCellCol);
    ftDefault:
      ExtractFromDefault;
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

  @param(AChangedItems  Set with elements identifying whether workbook, worksheet
                        cell or selection has changed.)
  @param(AData          If AChangedItems contains nliCell then AData points to
                        the modified cell.)
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

  case FFormatTarget of
    ftCell:
      begin
        activeCell := GetActiveCell;
        if (([lniCell]*AChangedItems <> []) and (PCell(AData) = activeCell)) or
           (lniSelection in AChangedItems)
        then
          ExtractFromCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
      end;
    ftRow:
      if (([lniRow] * AChangedItems <> []) and ({%H-}PtrUInt(AData) = Worksheet.ActiveCellRow)) or
         (lniSelection in AChangedItems)
      then
        ExtractFromRow(Worksheet.ActiveCellRow);
    ftCol:
      if (([lniCol] * AChangedItems <> []) and ({%H-}PtrUInt(AData) = Worksheet.ActiveCellCol)) or
         (lniSelection in AChangedItems)
      then
        ExtractFromCol(Worksheet.ActiveCellCol);
    ftDefault:
      ExtractFromDefault;
  end;
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
        if Assigned(FOnAddColors) then
          FOnAddColors(self)
        else begin
          // By default, add the Excel2 colors.
          AddColor(scBlack, GetColorName(scBlack));
          AddColor(scWhite, GetColorName(scWhite));
          AddColor(scRed, GetColorName(scRed));
          AddColor(scGreen, GetColorName(scGreen));
          AddColor(scBlue, GetColorName(scBlue));
          AddColor(scYellow, GetColorName(scYellow));
          AddColor(scMagenta, GetColorName(scMagenta));
          AddColor(scCyan, GetColorName(scCyan));
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
        ApplyFormat(r, c);
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
    inherited Style := csOwnerDrawFixed
  else
    inherited Style := csDropdown;

  Populate;
  ExtractFromSheet;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the FormatTarget
-------------------------------------------------------------------------------}
procedure TsCellCombobox.SetFormatTarget(AValue: TsFormatTarget);
begin
  if AValue = FFormatTarget then
    exit;
  FFormatTarget := AValue;
  ExtractFromSheet;
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
  FExpanded := [ienFormatSettings, ienConditionalFormats, ienPageLayout,
    ienFonts, ienFormats, ienEmbeddedObj, ienImages, ienCryptoInfo];
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
  if (pos('Conditional formats', s) > 0) or (pos('ConditionalFormats', s) > 0) then
  begin
    if (ienConditionalFormats in expNodes)
      then Exclude(expNodes, ienConditionalFormats)
      else Include(expNodes, ienConditionalFormats);
  end else
  if (pos('Page layout', s) > 0) or (pos('PageLayout', s) > 0) then
  begin
    if (ienPageLayout in expNodes)
      then Exclude(expNodes, ienPageLayout)
      else Include(expNodes, ienPageLayout);
  end else
  if (pos('Images', s) > 0) then
  begin
    if (ienEmbeddedObj in expNodes)
      then Exclude(expNodes, ienEmbeddedObj)
      else Include(expNodes, ienEmbeddedObj);
    if (ienImages in expNodes)
      then Exclude(expNodes, ienImages)
      else Include(expNodes, ienImages);
  end else
  if (pos('Fonts', s) > 0) then
  begin
    if (ienFonts in expNodes)
      then Exclude(expNodes, ienFonts)
      else Include(expNodes, ienFonts);
  end else
  if (pos('Cell formats', s) > 0) then
  begin
    if (ienFormats in expNodes)
      then Exclude(expNodes, ienFormats)
      else Include(expNodes, ienFormats);
  end else
  if (pos('CryptoInfo', s) > 0) then
  begin
    if (ienCryptoInfo in expNodes)
      then Exclude(expNodes, ienCryptoInfo)
      else Include(expNodes, ienCryptoInfo);
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
    if sheet <> nil then begin
      FCurrRow := sheet.ActiveCellRow;
      FCurrCol := sheet.ActiveCellCol;
      cell := sheet.FindCell(FCurrRow, FCurrCol);
    end;
  end;

  list := TStringList.Create;
  try
    case FMode of
      imCellValue      : UpdateCellValue(cell, list);
      imCellProperties : UpdateCellProperties(cell, list);
      imWorksheet      : UpdateWorksheet(sheet, list);
      imWorkbook       : UpdateWorkbook(book, list);
      imRow            : UpdateRow(FCurrRow, list);
      imCol            : UpdateCol(FCurrCol, list);
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

  @param(AChangedItems  Set with elements identifying whether workbook, worksheet
                        cell or selection has changed.)
  @param(AData          If AChangedItems contains nliCell then AData points to
                        the modified cell.)
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
begin
  case FMode of
    imWorkbook:
      if ([lniWorkbook, lniWorksheet]*AChangedItems <> []) then
        DoUpdate;
    imWorksheet:
      if ([lniWorksheet, lniSelection, lniWorksheetZoom]*AChangedItems <> []) then
        DoUpdate;
    imCellValue, imCellProperties:
      if ([lniCell, lniSelection]*AChangedItems <> []) then
        DoUpdate;
    imRow:
      begin
        if ([lniSelection] * AChangedItems <> []) then begin
          if AData <> nil then
            FCurrRow := PCell(AData)^.Row;
        end else if ([lniRow] * AChangedItems <> []) then
          FCurrRow := {%H-}PtrInt(AData)
        else
          exit;
        DoUpdate;
      end;
    imCol:
      begin
        if ([lniSelection] * AChangedItems <> []) then begin
          if AData <> nil then
            FCurrCol := PCell(AData)^.Col;
        end else if ([lniCol] * AChangedItems <> []) then
          FCurrCol := {%H-}PtrInt(AData)
        else
          exit;
        DoUpdate;
      end;
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
  r1, r2, c1, c2: Cardinal;
  rtp: TsRichTextParam;
begin
  if ACell <> nil then
    UpdateFormatProperties(ACell^.FormatIndex, AStrings)
  else
    UpdateFormatProperties(-1, AStrings);

  if (ACell <> nil) and (Length(ACell^.RichTextParams) > 0) then
  begin
    s := '';
    for rtp in ACell^.RichTextParams do
      s := Format('%s; Font #%d at pos %d', [s, rtp.FontIndex, rtp.FirstIndex]);
    Delete(s, 1, 2);
    if s = '' then s := '(none)';
    AStrings.Add('Rich-text parameters='+s);
  end else
    AStrings.Add('Rich-text parameters=(none)');

  if (Worksheet = nil) or not Worksheet.IsMerged(ACell) then
    AStrings.Add('Merged range=(none)')
  else
  begin
    Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    AStrings.Add('Merged range=' + GetCellRangeString(r1, c1, r2, c2));
  end;
end;

procedure TsSpreadsheetInspector.UpdateFormatProperties(AFormatIndex: integer;
  AStrings: TStrings);
var
  s: String;
  cb: TsCellBorder;
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams;
  cp: TsCellProtection;
begin
  if AFormatIndex > -1 then
    fmt := Workbook.GetCellFormat(AFormatIndex)
  else
    InitFormatRecord(fmt);

  if (AFormatIndex = -1)
    then AStrings.Add('FormatIndex=(default)')
    else AStrings.Add(Format('FormatIndex=%d', [AFormatIndex]));

  if (AFormatIndex = -1) or not (uffFont in fmt.UsedFormattingFields)
    then AStrings.Add('FontIndex=(default)')
    else AStrings.Add(Format('FontIndex=%d (%s)', [
           fmt.FontIndex,
           Workbook.GetFontAsString(fmt.FontIndex)
         ]));

  if (AFormatIndex = -1) or not (uffTextRotation in fmt.UsedFormattingFields)
    then AStrings.Add('TextRotation=(default)')
    else AStrings.Add(Format('TextRotation=%s', [
           GetEnumName(TypeInfo(TsTextRotation), ord(fmt.TextRotation))
         ]));

  if (AFormatIndex = -1) or not (uffHorAlign in fmt.UsedFormattingFields)
    then AStrings.Add('HorAlignment=(default)')
    else AStrings.Add(Format('HorAlignment=%s', [
           GetEnumName(TypeInfo(TsHorAlignment), ord(fmt.HorAlignment))
         ]));

  if (AFormatIndex = -1) or not (uffVertAlign in fmt.UsedFormattingFields)
    then AStrings.Add('VertAlignment=(default)')
    else AStrings.Add(Format('VertAlignment=%s', [
           GetEnumName(TypeInfo(TsVertAlignment), ord(fmt.VertAlignment))
         ]));

  if (AFormatIndex = -1) or not (uffWordwrap in fmt.UsedFormattingFields)
    then AStrings.Add('Wordwrap=(default)')
    else AStrings.Add(Format('Wordwrap=%s', [
           BoolToStr(uffWordwrap in fmt.UsedFormattingFields, true)
        ]));

  if (AFormatIndex = -1) or not (uffBorder in fmt.UsedFormattingFields) then
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
    if AFormatIndex = -1 then
      AStrings.Add(Format('BorderStyles[%s]=(default)', [
        GetEnumName(TypeInfo(TsCellBorder), ord(cb))]))
    else
      AStrings.Add(Format('BorderStyles[%s]=%s, %s', [
        GetEnumName(TypeInfo(TsCellBorder), ord(cb)),
        GetEnumName(TypeInfo(TsLineStyle), ord(fmt.BorderStyles[cb].LineStyle)),
        GetColorName(fmt.BorderStyles[cb].Color)]));

  if (AFormatIndex = -1) or not (uffBackground in fmt.UsedformattingFields) then
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

  if (AFormatIndex = -1) or not (uffNumberFormat in fmt.UsedFormattingFields) then
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

  if (AFormatIndex = -1) or not (uffBiDi in fmt.UsedFormattingFields) then
    AStrings.Add('BiDi=(bdDefault)')
  else
    AStrings.Add(Format('BiDiMode=%s', [
      GetEnumName(TypeInfo(TsBiDiMode), ord(fmt.BiDiMode))]));

  if (AFormatIndex = -1) then
    AStrings.Add('Protection=(default)')
  else begin
    if Worksheet.IsProtected and (spCells in Worksheet.Protection) then begin
      s := '';
      for cp in TsCellProtection do
        if cp in fmt.Protection then
          s := s + ', ' + GetEnumName(TypeInfo(TsCellProtection), ord(cp));
      if s <> '' then Delete(s, 1, 2) else s := '(not protected)';
    end else
      s := '(not protected)';
    AStrings.Add('Protection=' + s);
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
  Creates a string list containing the properties of a column.
  The string list items are name-value pairs in the format "name=value".
  The string list is displayed in the inspector's grid.

  @param  ACol       index of the investigated column
  @param  AStrings   Stringlist receiving the name-value pairs.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.UpdateCol(ACol: Integer; AStrings: TStrings);
var
  unitStr: String;
  lCol: PCol;
begin
  if (Workbook = nil) or (Worksheet = nil) then
    exit;
  if (ACol < 0) or (ACol <> Integer(Worksheet.ActiveCellCol)) then
    exit;

  unitStr := SizeUnitNames[Workbook.Units];
  lCol := Worksheet.FindCol(ACol);
  AStrings.Add(Format('Col=%d', [ACol]));
  if lCol <> nil then
  begin
    AStrings.Add(Format('Width=%.1f %s (%.1f pt)', [
      lCol^.Width, unitstr, Workbook.ConvertUnits(lCol^.Width, Workbook.Units, suPoints)
    ]));
    AStrings.Add(Format('ColWidthType=%s', [
      ColWidthTypeNames[lCol^.ColWidthType]
    ]));
    UpdateFormatProperties(lCol^.FormatIndex, AStrings);
    AStrings.Add(Format('Hidden=%s', [
      BoolToStr(croHidden in lCol^.Options, true)
    ]));
    AStrings.Add(Format('PageBreak=%s', [
      BoolToStr(croPageBreak in lCol^.Options, true)
    ]));
  end else
  begin
    AStrings.Add('No column record=');
    AStrings.Add(Format('DefaultColWidth=%.1f %s (%.1f pt)', [
      Worksheet.ReadDefaultColWidth(Workbook.Units), unitStr,
      Worksheet.ReadDefaultColWidth(suPoints)
    ]));
  //  UpdateFormatProperties(-1, AStrings);
  end;
end;


{@@ ----------------------------------------------------------------------------
  Creates a string list containing the properties of a row.
  The string list items are name-value pairs in the format "name=value".
  The string list is displayed in the inspector's grid.

  @param  ARow       index of the investigated row
  @param  AStrings   Stringlist receiving the name-value pairs.
-------------------------------------------------------------------------------}
procedure TsSpreadsheetInspector.UpdateRow(ARow: Integer; AStrings: TStrings);
var
  lRow: PRow;
  unitStr: String;
begin
  if (Workbook = nil) or (Worksheet = nil) then
    exit;
  if (ARow < 0) or (ARow <> Integer(Worksheet.ActiveCellRow)) then
    exit;

  unitStr := SizeUnitNames[Workbook.Units];
  lRow := Worksheet.FindRow(ARow);
  AStrings.Add(Format('Row=%d', [ARow]));
  if lRow <> nil then
  begin
    AStrings.Add(Format('Height=%.1f %s (%.1f pt)', [
      lRow^.Height, unitStr, Workbook.ConvertUnits(lRow^.Height, Workbook.Units, suPoints)
    ]));
    AStrings.Add(Format('RowHeightType=%s', [
      RowHeightTypeNames[lRow^.RowHeightType]
    ]));
    UpdateFormatProperties(lRow^.FormatIndex, AStrings);
    AStrings.Add(Format('Hidden=%s', [
      BoolToStr(croHidden in lRow^.Options, true)
    ]));
    AStrings.Add(Format('PageBreak=%s', [
      BoolToStr(croPageBreak in lRow^.Options, true)
    ]));
  end else
  begin
    AStrings.Add('No row record=');
    AStrings.Add(Format('DefaultRowHeight=%.1f %s (%.1f pt)', [
      Worksheet.ReadDefaultRowHeight(Workbook.Units), unitStr,
      Worksheet.ReadDefaultRowHeight(suPoints)
    ]));
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
  embobj: TsEmbeddedObj;
  bp: TsWorkbookProtection;
begin
  if AWorkbook = nil then
  begin
    AStrings.Add('FileName=');
    AStrings.Add('FileFormat=');
    AStrings.Add('Options=');
    AStrings.Add('ActiveWorksheet=');
    AStrings.Add('FormatSettings=');
    AStrings.Add('Images=');
    AStrings.Add('Protection=');
  end else
  begin
    AStrings.Add(Format('FileName=%s', [AWorkbook.FileName]));
    if AWorkbook.FileFormatID = -1 then
      AStrings.Add('FileFormat=(unknown)')
    else
      AStrings.Add(Format('FileFormat=%d [%s]', [
        AWorkbook.FileFormatID, GetSpreadTechnicalName(AWorkbook.FileFormatID)
      ]));
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

    if (ienEmbeddedObj in FExpanded) then begin
      AStrings.Add('(-) Images=');
      for i:=0 to AWorkbook.GetEmbeddedObjCount-1 do
      begin
        embObj := AWorkbook.GetEmbeddedObj(i);
        AStrings.Add('  Filename='+embobj.FileName);
        AStrings.Add(Format('  ImageWidth=%.2f mm', [embObj.ImageWidth]));
        AStrings.Add(Format('  ImageHeight=%.2f mm', [embObj.ImageHeight]));
      end;
    end else
      AStrings.Add('(+) Images=(dblclick for more...)');

    if (ienFonts in FExpanded) then begin
      AStrings.Add('(-) Fonts=');
      for i:=0 to AWorkbook.GetFontCount-1 do
        AStrings.Add(Format('  Font%d=%s', [i, AWorkbook.GetFontAsString(i)]));
    end else
      AStrings.Add('(+) Fonts=(dblclick for more...)');

    if (ienFormats in FExpanded) then begin
      AStrings.Add('(-) Cell formats=');
      for i:=0 to AWorkbook.GetNumCellFormats-1 do
        AStrings.Add(Format('  CellFormat%d=%s', [i, AWorkbook.GetCellFormatAsString(i)]));
    end else
      AStrings.Add('(+) Cell formats=(dblclick for more...)');

    s := '';
    for bp in TsWorkbookProtection do
      if bp in AWorkbook.Protection then
        s := s + ', ' + GetEnumName(TypeInfo(TsWorkbookProtection), ord(bp));
    if s <> '' then Delete(s, 1, 2) else s := '(default)';
    AStrings.Add('Protection=' + s);

    if (ienCryptoInfo in FExpanded) then begin
      AStrings.Add('(-) CryptoInfo=');
      AStrings.Add(Format('  PasswordHash=%s', [Workbook.CryptoInfo.PasswordHash]));
      AStrings.Add(Format('  Algorithm=%s', [AlgorithmToStr(Workbook.CryptoInfo.Algorithm, auExcel)]));
      AStrings.Add(Format('  SaltValue=%s', [Workbook.CryptoInfo.SaltValue]));
      AStrings.Add(Format('  SpinCount=%d', [Workbook.CryptoInfo.SpinCount]));
    end else
      AStrings.Add('(+) CryptoInfo=(dblclick for more...)');

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
  i: Integer;
  s: String;
  po: TsPrintOption;
  img: TsImage;
  embObj: TsEmbeddedObj;
  so: TsSheetOption;
  sp: TsWorksheetProtection;
begin
  if ASheet = nil then
  begin
    AStrings.Add('Name=');
    AStrings.Add('Index=');
    AStrings.Add('First row=');
    AStrings.Add('Last row=');
    AStrings.Add('First column=');
    AStrings.Add('Last column=');
    AStrings.Add('Active cell=');
    AStrings.Add('Selection=');
    AStrings.Add('Default column width=');
    AStrings.Add('Default row height=');
    AStrings.Add('Zoom factor=');
    AStrings.Add('Page layout=');
    AStrings.Add('Options=');
    AStrings.Add('Protection=');
    AStrings.Add('TabColor=');
    AStrings.Add('Conditional formats=');
  end else
  begin
    AStrings.Add(Format('Name=%s', [ASheet.Name]));
    AStrings.Add(Format('Index=%d', [ASheet.Index]));
    AStrings.Add(Format('First row=%d', [Integer(ASheet.GetFirstRowIndex)]));
    AStrings.Add(Format('Last row=%d', [ASheet.GetLastRowIndex(true)]));
    AStrings.Add(Format('First column=%d', [Integer(ASheet.GetFirstColIndex)]));
    AStrings.Add(Format('Last column=%d', [ASheet.GetLastColIndex(true)]));
    AStrings.Add(Format('Active cell=%s',
      [GetCellString(ASheet.ActiveCellRow, ASheet.ActiveCellCol)]));
    AStrings.Add(Format('Selection=%s', [ASheet.GetSelectionAsString]));
    AStrings.Add(Format('Default column width=%.1f %s', [
      ASheet.ReadDefaultColWidth(ASheet.Workbook.Units),
      SizeUnitNames[ASheet.Workbook.Units]]));
    AStrings.Add(Format('Default row height=%.1f %s', [
      ASheet.ReadDefaultRowHeight(ASheet.Workbook.Units),
      SizeUnitNames[ASheet.Workbook.Units]]));
    AStrings.Add(Format('Zoom factor=%d%%', [round(ASheet.ZoomFactor*100)]));
    AStrings.Add(Format('Comments=%d items', [ASheet.Comments.Count]));
    AStrings.Add(Format('Hyperlinks=%d items', [ASheet.Hyperlinks.Count]));
    AStrings.Add(Format('MergedCells=%d items', [ASheet.MergedCells.Count]));
    AStrings.Add(Format('TabColor=$%.8x (%s)', [ASheet.TabColor, GetColorName(ASheet.TabColor)]));
                                    (*
    if ienConditionalFormats in FExpanded then
    begin
      AStrings.Add('(-) Conditional formats=');
      AStrings.Add(Format('  Count=%d', [ASheet.ConditionalFormatCount]));
      for i := 0 to ASheet.ConditionalFormatCount-1 do
      begin
        cf := ASheet.ReadConditionalFormat(i);
        AStrings.Add('  Item #' + IntToStr(i) + ':');
        with cf.CellRange do
          AStrings.Add(Format('    CellRange=%s', [GetCellRangeString(Row1, Col1, Row2, Col2)]));
      end;
    end else
    begin
      AStrings.Add('(+) Conditional formats=(dblclick for more...)');
    end;
                       *)
    if ienPageLayout in FExpanded then
    begin
      AStrings.Add('(-) Page layout=');
      AStrings.Add(Format('  Orientation=%s', [
        GetEnumName(TypeInfo(TsPageOrientation),
        ord(ASheet.PageLayout.Orientation))]));
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
        AStrings.Add('  Start page number=automatic');
      AStrings.Add(Format('  Scaling factor (Zoom)=%d%%',
        [ASheet.PageLayout.ScalingFactor]));
      AStrings.Add(Format('  Copies=%d', [ASheet.PageLayout.Copies]));
      if (ASheet.PageLayout.Options * [poDifferentOddEven, poDifferentFirst] <> []) then
      begin
        AStrings.Add(Format('  Header (first)=%s',
          [StringReplace(ASheet.PageLayout.Headers[0], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Header (odd)=%s',
          [StringReplace(ASheet.PageLayout.Headers[1], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Header (even)=%s',
          [StringReplace(ASheet.PageLayout.Headers[2], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Footer (first)=%s',
          [StringReplace(ASheet.PageLayout.Footers[0], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Footer (odd)=%s',
          [StringReplace(ASheet.PageLayout.Footers[1], LineEnding, '\n', [rfReplaceall])]));
        AStrings.Add(Format('  Footer (even)=%s',
          [StringReplace(ASheet.PageLayout.Footers[2], LineEnding, '\n', [rfReplaceAll])]));
      end else
      begin
        AStrings.Add(Format('  Header=%s', [StringReplace(ASheet.PageLayout.Headers[1], LineEnding, '\n', [rfReplaceAll])]));
        AStrings.Add(Format('  Footer=%s', [StringReplace(ASheet.PageLayout.Footers[1], LineEnding, '\n', [rfReplaceAll])]));
      end;

      if ASheet.PageLayout.HeaderImages[hfsLeft].Index > -1 then
        AStrings.Add(Format('  HeaderImage, left=%s',
          [ASheet.Workbook.GetEmbeddedObj(ASheet.PageLayout.HeaderImages[hfsLeft].Index).FileName]))
      else
        AStrings.Add('  HeaderImage, left =');
      if ASheet.PageLayout.HeaderImages[hfsCenter].Index > -1 then
        AStrings.Add(Format('  HeaderImage, center=%s',
          [ASheet.Workbook.GetEmbeddedObj(ASheet.PageLayout.HeaderImages[hfsCenter].Index).FileName]))
      else
        AStrings.Add('  HeaderImage, center=');
      if ASheet.PageLayout.HeaderImages[hfsRight].Index > -1 then
        AStrings.Add(Format('  HeaderImage, right=%s',
          [ASheet.Workbook.GetEmbeddedObj(ASheet.PageLayout.HeaderImages[hfsRight].Index).FileName]))
      else
        AStrings.Add('  HeaderImage, right=');

      if ASheet.PageLayout.FooterImages[hfsLeft].Index > -1 then
        AStrings.Add(Format('  FooterImage, left=%s',
          [ASheet.Workbook.GetEmbeddedObj(ASheet.PageLayout.FooterImages[hfsLeft].Index).FileName]))
      else
        AStrings.Add('  FooterImage, left =');
      if ASheet.PageLayout.FooterImages[hfsCenter].Index > -1 then
        AStrings.Add(Format('  FooterImage, center=%s',
          [ASheet.Workbook.GetEmbeddedObj(ASheet.PageLayout.FooterImages[hfsCenter].Index).FileName]))
      else
        AStrings.Add('  FooterImage, center=');
      if ASheet.PageLayout.FooterImages[hfsRight].Index > -1 then
        AStrings.Add(Format('  FooterImage, right=%s', [
          ASheet.Workbook.GetEmbeddedObj(ASheet.PageLayout.FooterImages[hfsRight].Index).FileName]))
      else
        AStrings.Add('  FooterImage, right=');

      if ASheet.PageLayout.NumPrintRanges = 0 then
        AStrings.Add('  Print ranges=')
      else
        for i := 0 to ASheet.PageLayout.NumPrintRanges-1 do
          with ASheet.PageLayout.PrintRange[i] do
            AStrings.Add(Format('  Print range #%d=$%s$%s:$%s$%s', [ i,
              GetColString(Col1), GetRowString(Row1), GetColString(Col2), GetRowString(Row2)
            ]));

      if ASheet.PageLayout.RepeatedRows.FirstIndex = UNASSIGNED_ROW_COL_INDEX then
        AStrings.Add('  Repeated rows=')
      else
      if ASheet.PageLayout.RepeatedRows.FirstIndex = ASheet.PageLayout.RepeatedRows.LastIndex then
        AStrings.Add(Format('  Repeated rows=$%s', [
          GetRowString(ASheet.PageLayout.RepeatedRows.FirstIndex)
        ]))
      else
        AStrings.Add(Format('  Repeated rows=$%s:$%s', [
          GetRowString(ASheet.PageLayout.RepeatedRows.FirstIndex),
          GetRowString(ASheet.PageLayout.RepeatedRows.lastIndex)
        ]));

      if ASheet.PageLayout.RepeatedCols.FirstIndex = UNASSIGNED_ROW_COL_INDEX then
        AStrings.Add('  Repeated columns=')
      else
      if ASheet.PageLayout.RepeatedCols.FirstIndex = ASheet.PageLayout.RepeatedCols.LastIndex then
        AStrings.Add(Format('  Repeated columns=$%s', [
          GetColString(ASheet.PageLayout.RepeatedCols.FirstIndex)
        ]))
      else
        AStrings.Add(Format('  Repeated columns=$%s:$%s', [
          GetColString(ASheet.PageLayout.RepeatedCols.FirstIndex),
          GetColString(ASheet.PageLayout.RepeatedCols.lastIndex)
        ]));

      s := '';
      for po in TsPrintOption do
        if po in ASheet.PageLayout.Options then s := s + '; ' + GetEnumName(typeInfo(TsPrintOption), ord(po));
      if s <> '' then Delete(s, 1, 2);
      AStrings.Add(Format('  Options=%s', [s]));
    end else
      AStrings.Add('(+) Page layout=(dblclick for more...)');

    if (ienImages in FExpanded) then begin
      AStrings.Add('(-) Images=');
      for i:=0 to ASheet.GetImageCount-1 do
      begin
        img := ASheet.GetImage(i);
        AStrings.Add(Format('  Row=%d', [img.Row]));
        AStrings.Add(Format('  Col=%d', [img.Col]));
        embObj := ASheet.Workbook.GetEmbeddedObj(img.Index);
        AStrings.Add(Format('  Index=%d [%s; %.2fmm x %.2fmm]',
          [img.Index, embobj.FileName, embObj.ImageWidth, embObj.ImageHeight]));
        AStrings.Add(Format('  OffsetX=%.2f mm', [img.OffsetX]));
        AStrings.Add(Format('  OffsetY=%.2f mm', [img.OffsetY]));
        AStrings.Add(Format('  ScaleX=%.2f', [img.ScaleX]));
        AStrings.Add(Format('  ScaleY=%.2f', [img.ScaleY]));
        AStrings.Add(Format('  HyperlinkTarget=%s', [img.HyperlinkTarget]));
        AStrings.Add(Format('  HyperlinkTooltip=%s', [img.HyperlinkToolTip]));
      end;
    end else
      AStrings.Add('(+) Images=(dblclick for more...)');

    s := '';
    for so in TsSheetOption do
      if so in ASheet.Options then
        s := s + ', ' + GetEnumName(TypeInfo(TsSheetOption), ord(so));
    if s <> '' then Delete(s, 1, 2);
    AStrings.Add('Options='+s);

    if ASheet.IsProtected then begin
      s := '';
      for sp in TsWorksheetProtection do
        if sp in ASheet.Protection then
          s := s + ', ' + GetEnumName(TypeInfo(TsWorksheetProtection), ord(sp));
      if s <> '' then Delete(s, 1, 2) else s := '(default)';
    end else
      s := '(not protected)';
    AStrings.Add('Protection=' + s);

    if (ienCryptoInfo in FExpanded) then begin
      AStrings.Add('(-) CryptoInfo=');
      AStrings.Add(Format('  PasswordHash=%s', [Worksheet.CryptoInfo.PasswordHash]));
      AStrings.Add(Format('  Algorithm=%s', [AlgorithmToStr(Worksheet.CryptoInfo.Algorithm, auExcel)]));
      AStrings.Add(Format('  SaltValue=%s', [Worksheet.CryptoInfo.SaltValue]));
      AStrings.Add(Format('  SpinCount=%d', [Worksheet.CryptoInfo.SpinCount]));
    end else
      AStrings.Add('(+) CryptoInfo=(dblclick for more...)');

  end;
end;


initialization
  RegisterPropertyToSkip(TsSpreadsheetInspector, 'RowHeights',
    'For compatibility with older Laz versions.', '');

  RegisterPropertyToSkip(TsSpreadsheetInspector, 'ColWidths',
    'For compatibility with older Laz versions.', '');

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

end.
