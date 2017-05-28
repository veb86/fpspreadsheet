{@@ ----------------------------------------------------------------------------
  Unit fpspreadsheet implements a <b>grid</b> component which can load and
  write data from/to FPSpreadsheet documents.

  Can either be used alone or in combination with a TsWorkbookSource component.
  The latter method requires less written code.

  AUTHORS: Felipe Monteiro de Carvalho, Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpspreadsheetgrid;

{$mode objfpc}{$H+}
{$I ..\fps.inc}

{.$DEFINE GRID_DEBUG}

{ To do:
 - When Lazarus 1.4 comes out remove the workaround for the RGB2HLS bug in
   FindNearestPaletteIndex.
 - Arial bold is not shown as such if loaded from ods
 - Background color of first cell is ignored. }

interface

uses
  Classes, SysUtils, LMessages, LResources, Variants,
  Forms, Controls, Graphics, Dialogs, Grids, StdCtrls, ExtCtrls,
  fpstypes, fpspreadsheet, fpspreadsheetctrls;

type

  { TsCustomWorksheetGrid }
  TsCustomWorksheetGrid = class;

  TsAutoExpandMode = (
    {@@ Expands grid dimensions if a cell is written outside current grid dimensions }
    aeData,
    {@@ Expands grid dimensions if navigation goes outside current grid dimensions }
    aeNavigation,
    {@@ Expands grid dimensions to DEFAULT_ROW_COUNT and DEFAULT_COL_COUNT }
    aeDefault
  );
  TsAutoExpandModes = set of TsAutoExpandMode;

  TsEditorLineMode = (elmSingleLine, elmMultiLine);

  TsHyperlinkClickEvent = procedure(Sender: TObject;
    const AHyperlink: TsHyperlink) of object;

  TsSelPen = class(TPen)
  public
    constructor Create; override;
  published
    property Width stored true default 3;
    property JoinStyle default pjsMiter;
  end;

  TMultilineStringCellEditor = class(TMemo)
  private
    FGrid: TsCustomWorksheetGrid;
    FCol,FRow:Integer;
  protected
    procedure WndProc(var TheMessage : TLMessage); override;
    procedure Change; override;
    procedure KeyDown(var Key : Word; Shift : TShiftState); override;
    procedure msg_SetValue(var Msg: TGridMessage); message GM_SETVALUE;
    procedure msg_GetValue(var Msg: TGridMessage); message GM_GETVALUE;
    procedure msg_SetGrid(var Msg: TGridMessage); message GM_SETGRID;
    procedure msg_SelectAll(var Msg: TGridMessage); message GM_SELECTALL;
    procedure msg_SetPos(var Msg: TGridMessage); message GM_SETPOS;
    procedure msg_GetGrid(var Msg: TGridMessage); message GM_GETGRID;
  public
    constructor Create(Aowner : TComponent); override;
    procedure EditingDone; override;
//    property EditText;
    property OnEditingDone;
  end;


//  TsSelectionRectMode = (srmDThickXOR, srmThick, srmDottedXOR,
  {@@ TsCustomWorksheetGrid is the ancestor of TsWorksheetGrid and is able to
    display spreadsheet data along with their formatting. }
  TsCustomWorksheetGrid = class(TCustomDrawGrid, IsSpreadsheetControl)
  private
    { Private declarations }
    FWorkbookSource: TsWorkbookSource;
    FInternalWorkbookSource: TsWorkbookSource;
    FHeaderCount: Integer;
    FFrozenCols: Integer;
    FFrozenRows: Integer;
    FEditText: String;
    FLockCount: Integer;
    FLockSetup: Integer;
    FEditing: Boolean;
    FCellFont: TFont;
    FAutoCalc: Boolean;
    FTextOverflow: Boolean;
    FReadFormulas: Boolean;
    FDrawingCell: PCell;
    FTextOverflowing: Boolean;
    FAutoExpand: TsAutoExpandModes;
    FEnhEditMode: Boolean;
    FSelPen: TsSelPen;
    FFrozenBorderPen: TPen;
    FHyperlinkTimer: TTimer;
    FHyperlinkCell: PCell;      // Selected cell if it stores a hyperlink
    FDefRowHeight100: Integer;  // Default row height for 100% zoom factor, in pixels
    FDefColWidth100: Integer;   // Default col width for 100% zoom factor, in pixels
    FZoomLock: Integer;
    FRowHeightLock: Integer;
    FActiveCellLock: Integer;
    FTopLeft: TPoint;
    FReadOnly: Boolean;
    FOnClickHyperlink: TsHyperlinkClickEvent;
    FOldEditorText: String;
    FMultilineStringEditor: TMultilineStringCellEditor;
    FLineMode: TsEditorLineMode;
    FAllowDragAndDrop: Boolean;
    FDragStartCol, FDragStartRow: Integer;
    FOldDragStartCol, FOldDragStartRow: Integer;
    FDragSelection: TGridRect;
    function CalcAutoRowHeight(ARow: Integer): Integer;
    function CalcColWidthFromSheet(AWidth: Single): Integer;
    function CalcRowHeightFromSheet(AHeight: Single): Integer;
    function CalcRowHeightToSheet(AHeight: Integer): Single;
    procedure ChangedCellHandler(ASender: TObject; ARow, ACol: Cardinal);
    procedure ChangedFontHandler(ASender: TObject; ARow, ACol: Cardinal);
    procedure FixNeighborCellBorders(ACell: PCell);
    function GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer;
      ACell: PCell; out ABorderStyle: TsCellBorderStyle): Boolean;

    // Setter/Getter
    function GetBackgroundColor(ACol, ARow: Integer): TsColor;
    function GetBackgroundColors(ALeft, ATop, ARight, ABottom: Integer): TsColor;
    function GetCellBiDiMode(ACol, ARow: Integer): TsBiDiMode;
    function GetCellBorder(ACol, ARow: Integer): TsCellBorders;
    function GetCellBorders(ALeft, ATop, ARight, ABottom: Integer): TsCellBorders;
    function GetCellBorderStyle(ACol, ARow: Integer; ABorder: TsCellBorder): TsCellBorderStyle;
    function GetCellBorderStyles(ALeft, ATop, ARight, ABottom: Integer;
      ABorder: TsCellBorder): TsCellBorderStyle;
    function GetCellComment(ACol, ARow: Integer): string;
    function GetCellFont(ACol, ARow: Integer): TFont;
    function GetCellFonts(ALeft, ATop, ARight, ABottom: Integer): TFont;
    function GetCellFontColor(ACol, ARow: Integer): TsColor;
    function GetCellFontColors(ALeft, ATop, ARight, ABottom: Integer): TsColor;
    function GetCellFontName(ACol, ARow: Integer): String;
    function GetCellFontNames(ALeft, ATop, ARight, ABottom: Integer): String;
    function GetCellFontSize(ACol, ARow: Integer): Single;
    function GetCellFontSizes(ALeft, ATop, ARight, ABottom: Integer): Single;
    function GetCellFontStyle(ACol, ARow: Integer): TsFontStyles;
    function GetCellFontStyles(ALeft, ATop, ARight, ABottom: Integer): TsFontStyles;
    function GetCellValue(ACol, ARow: Integer): variant;
    function GetColWidths(ACol: Integer): Integer;
    function GetDefColWidth: Integer;
    function GetDefRowHeight: Integer;
    function GetHorAlignment(ACol, ARow: Integer): TsHorAlignment;
    function GetHorAlignments(ALeft, ATop, ARight, ABottom: Integer): TsHorAlignment;
    function GetHyperlink(ACol, ARow: Integer): String;
    function GetNumberFormat(ACol, ARow: Integer): String;
    function GetNumberFormats(ALeft, ATop, ARight, ABottom: Integer): String;
    function GetRowHeights(ARow: Integer): Integer;
    function GetShowGridLines: Boolean;
    function GetShowHeaders: Boolean; inline;
    function GetTextRotation(ACol, ARow: Integer): TsTextRotation;
    function GetTextRotations(ALeft, ATop, ARight, ABottom: Integer): TsTextRotation;
    function GetVertAlignment(ACol, ARow: Integer): TsVertAlignment;
    function GetVertAlignments(ALeft, ATop, ARight, ABottom: Integer): TsVertAlignment;
    function GetWorkbook: TsWorkbook;
    function GetWorkbookSource: TsWorkbookSource;
    function GetWorksheet: TsWorksheet;
    function GetWordwrap(ACol, ARow: Integer): Boolean;
    function GetWordwraps(ALeft, ATop, ARight, ABottom: Integer): Boolean;
    function GetZoomFactor: Double;
    procedure SetAutoCalc(AValue: Boolean);
    procedure SetBackgroundColor(ACol, ARow: Integer; AValue: TsColor);
    procedure SetBackgroundColors(ALeft, ATop, ARight, ABottom: Integer; AValue: TsColor);
    procedure SetCellBiDiMode(ACol, ARow: Integer; AValue: TsBiDiMode);
    procedure SetCellBorder(ACol, ARow: Integer; AValue: TsCellBorders);
    procedure SetCellBorders(ALeft, ATop, ARight, ABottom: Integer; AValue: TsCellBorders);
    procedure SetCellBorderStyle(ACol, ARow: Integer; ABorder: TsCellBorder;
      AValue: TsCellBorderStyle);
    procedure SetCellBorderStyles(ALeft, ATop, ARight, ABottom: Integer;
      ABorder: TsCellBorder; AValue: TsCellBorderStyle);
    procedure SetCellComment(ACol, ARow: Integer; AValue: String);
    procedure SetCellFont(ACol, ARow: Integer; AValue: TFont);
    procedure SetCellFonts(ALeft, ATop, ARight, ABottom: Integer; AValue: TFont);
    procedure SetCellFontColor(ACol, ARow: Integer; AValue: TsColor);
    procedure SetCellFontColors(ALeft, ATop, ARight, ABottom: Integer; AValue: TsColor);
    procedure SetCellFontName(ACol, ARow: Integer; AValue: String);
    procedure SetCellFontNames(ALeft, ATop, ARight, ABottom: Integer; AValue: String);
    procedure SetCellFontSize(ACol, ARow: Integer; AValue: Single);
    procedure SetCellFontSizes(ALeft, ATop, ARight, ABottom: Integer; AValue: Single);
    procedure SetCellFontStyle(ACol, ARow: Integer; AValue: TsFontStyles);
    procedure SetCellFontStyles(ALeft, ATop, ARight, ABottom: Integer;
      AValue: TsFontStyles);
    procedure SetCellValue(ACol, ARow: Integer; AValue: variant);
    procedure SetColWidths(ACol: Integer; AValue: Integer);
    procedure SetDefColWidth(AValue: Integer);
    procedure SetDefRowHeight(AValue: Integer);
    procedure SetEditorLineMode(AValue: TsEditorLineMode);
    procedure SetFrozenCols(AValue: Integer);
    procedure SetFrozenBorderPen(AValue: TPen);
    procedure SetFrozenRows(AValue: Integer);
    procedure SetHorAlignment(ACol, ARow: Integer; AValue: TsHorAlignment);
    procedure SetHorAlignments(ALeft, ATop, ARight, ABottom: Integer;
      AValue: TsHorAlignment);
    procedure SetHyperlink(ACol, ARow: Integer; AValue: String);
    procedure SetNumberFormat(ACol, ARow: Integer; AValue: String);
    procedure SetNumberFormats(ALeft, ATop, ARight, ABottom: Integer; AValue: String);
    procedure SetReadFormulas(AValue: Boolean);
    procedure SetRowHeights(ARow: Integer; AValue: Integer);
    procedure SetSelPen(AValue: TsSelPen);
    procedure SetShowGridLines(AValue: Boolean);
    procedure SetShowHeaders(AValue: Boolean);
    procedure SetTextRotation(ACol, ARow: Integer; AValue: TsTextRotation);
    procedure SetTextRotations(ALeft, ATop, ARight, ABottom: Integer;
      AValue: TsTextRotation);
    procedure SetVertAlignment(ACol, ARow: Integer; AValue: TsVertAlignment);
    procedure SetVertAlignments(ALeft, ATop, ARight, ABottom: Integer;
      AValue: TsVertAlignment);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
    procedure SetWordwrap(ACol, ARow: Integer; AValue: boolean);
    procedure SetWordwraps(ALeft, ATop, ARight, ABottom: Integer; AValue: boolean);
    procedure SetZoomFactor(AValue: Double);

    procedure HyperlinkTimerElapsed(Sender: TObject);

  protected
    { Protected declarations }
    procedure AdaptToZoomFactor;
    procedure AutoAdjustColumn(ACol: Integer); override;
    procedure AutoAdjustRow(ARow: Integer); virtual;
    procedure AutoExpandToCol(ACol: Integer; AMode: TsAutoExpandMode);
    procedure AutoExpandToRow(ARow: Integer; AMode: TsAutoExpandMode);
    function CalcTopLeft(AHeaderOnly: Boolean): TPoint;
    function CalcWorksheetColWidth(AValue: Integer): Single;
    function CalcWorksheetRowHeight(AValue: Integer): Single;
    function CanEditShow: boolean; override;
    function CellOverflow(ACol, ARow: Integer; AState: TGridDrawState;
      out ACol1, ACol2: Integer; var ARect: TRect): Boolean;
    procedure ColRowMoved(IsColumn: Boolean; FromIndex,ToIndex: Integer); override;
    procedure CreateHandle; override;
    procedure CreateNewWorkbook;
    procedure DblClick; override;
    procedure DefineProperties(Filer: TFiler); override;
    procedure DoCopyToClipboard; override;
    procedure DoCutToClipboard; override;
    procedure DoEditorShow; override;
    procedure DoPasteFromClipboard; override;
    procedure DoOnResize; override;
    procedure DoPrepareCanvas(ACol, ARow: Integer; AState: TGridDrawState); override;
    procedure DragOver(ASource: TObject; X, Y: Integer; AState: TDragState;
      var Accept: Boolean); override;
    procedure DrawAllRows; override;
    procedure DrawCellBorders(AGridPart: Integer = 0); overload;
    procedure DrawCellBorders(ACol, ARow: Integer; ARect: TRect; ACell: PCell); overload;
    procedure DrawCellGrid(ACol,ARow: Integer; ARect: TRect; AState: TGridDrawState); override;
    procedure DrawCommentMarker(ARect: TRect);
    procedure DrawDragSelection;
    procedure DrawFocusRect(aCol,aRow:Integer; ARect:TRect); override;
    procedure DrawFrozenPaneBorder(AStart, AEnd, ACoord: Integer; IsHor: Boolean);
    procedure DrawFrozenPanes;
    procedure DrawImages(AGridPart: Integer = 0);
    procedure DrawRow(aRow: Integer); override;
    procedure DrawSelection;
    procedure DrawTextInCell(ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState); override;
    procedure ExecuteHyperlink;
    procedure GenericPenChangeHandler(Sender: TObject);
    function GetCellHeight(ACol, ARow: Integer): Integer;
    function GetCellHintText(ACol, ARow: Integer): String; override;
    function GetCells(ACol, ARow: Integer): String; override;
    function GetCellText(ACol, ARow: Integer; ATrim: Boolean = true): String;
    function GetEditText(ACol, ARow: Integer): String; override;
    function GetDefaultHeaderColWidth: Integer;
    function HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
    procedure HeaderSizing(const IsColumn:boolean; const AIndex,ASize:Integer); override;
    procedure HeaderSized(IsColumn: Boolean; AIndex: Integer); override;
    procedure InternalDrawCell(ACol, ARow: Integer; AClipRect, ACellRect: TRect;
      AState: TGridDrawState);
    procedure InternalDrawRow(ARow, AFirstCol, ALastCol: Integer;
      AClipRect: TRect);
    procedure InternalDrawSelection(ASel: TGridRect; IsNormalSelection: boolean);
    procedure InternalDrawTextInCell(AText: String; ARect: TRect;
      ACellHorAlign: TsHorAlignment; ACellVertAlign: TsVertAlignment;
      ATextRot: TsTextRotation; ATextWrap: Boolean; AFontIndex: Integer;
      AOverrideTextColor: TColor; ARichTextParams: TsRichTextParams;
      AIsRightToLeft: Boolean);
    procedure KeyDown(var Key : Word; Shift : TShiftState); override;
    procedure MouseDown(Button: TMouseButton; Shift:TShiftState; X,Y:Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    function MouseOnCellBorder(const APoint: TPoint;
      const ACellRect: TGridRect): Boolean;
    procedure MouseUp(Button: TMouseButton; Shift:TShiftState; X,Y:Integer); override;
    procedure MoveSelection; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure PrepareCanvasFont;
    function RelaxAutoExpand: TsAutoExpandModes;
    procedure RestoreAutoExpand(AValue: TsAutoExpandModes);
    function  SelectCell(ACol, ARow: Integer): Boolean; override;
    procedure SetEditText(ACol, ARow: Longint; const AValue: string); override;
    procedure Setup;
    procedure Sort(AColSorting: Boolean; AIndex, AIndxFrom, AIndxTo:Integer); override;
    function ToPixels(AValue: Double): Integer;
    procedure TopLeftChanged; override;
    function TrimToCell(ACell: PCell): String;
    procedure WMHScroll(var message : TLMHScroll); message LM_HSCROLL;
    procedure WMVScroll(var message : TLMVScroll); message LM_VSCROLL;

    {@@ Allow built-in drag and drop }
    property AllowDragAndDrop: Boolean read FAllowDragAndDrop
      write FAllowDragAndDrop default true;
    {@@ Automatically recalculate formulas whenever a cell value changes. }
    property AutoCalc: Boolean read FAutoCalc write SetAutoCalc default false;
    {@@ Automatically expand grid dimensions }
    property AutoExpand: TsAutoExpandModes read FAutoExpand write FAutoExpand
      default [aeData, aeNavigation, aeDefault];
    {@@ Displays column and row headers in the fixed col/row style of the grid.
        Deprecated. Use ShowHeaders instead. }
    property DisplayFixedColRow: Boolean read GetShowHeaders write SetShowHeaders
      default true;
    {@@ Determines whether a single- or multi-line cell editor is used }
    property EditorLineMode: TsEditorLineMode read FLineMode write SetEditorLineMode
      default elmSingleLine;
    {@@ This number of columns at the left is "frozen", i.e. it is not possible to
        scroll these columns }
    property FrozenCols: Integer read FFrozenCols write SetFrozenCols;
    {@@ Pen for the line separating the frozen cols/rows from the regular grid }
    property FrozenBorderPen: TPen read FFrozenBorderPen write SetFrozenBorderPen;
    {@@ This number of rows at the top is "frozen", i.e. it is not possible to
        scroll these rows. }
    property FrozenRows: Integer read FFrozenRows write SetFrozenRows;
    {@@ Activates reading of RPN formulas. Should be turned off when
        non-implemented formulas crashe reading of the spreadsheet file. }
    property ReadFormulas: Boolean read FReadFormulas write SetReadFormulas;
    {@@ Pen used for drawing the selection rectangle }
    property SelectionPen: TsSelPen read FSelPen write SetSelPen;
    {@@ Shows/hides vertical and horizontal grid lines }
    property ShowGridLines: Boolean read GetShowGridLines write SetShowGridLines default true;
    {@@ Shows/hides column and row headers in the fixed col/row style of the grid. }
    property ShowHeaders: Boolean read GetShowHeaders write SetShowHeaders default true;
    {@@ Activates text overflow (cells reaching into neighbors) }
    property TextOverflow: Boolean read FTextOverflow write FTextOverflow default false;
    {@@ Event called when an external hyperlink is clicked }
    property OnClickHyperlink: TsHyperlinkClickEvent read FOnClickHyperlink write FOnClickHyperlink;

  public
    { public methods }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure AutoColWidth(ACol: Integer);
    procedure AutoRowHeight(ARow: Integer);
    procedure BeginUpdate;
    function CellRect(ACol1, ARow1, ACol2, ARow2: Integer): TRect; overload;
    procedure Clear;
    procedure DefaultDrawCell(ACol, ARow: Integer; var ARect: TRect;
      AState: TGridDrawState); override;
    procedure DeleteCol(AGridCol: Integer); reintroduce;
    procedure DeleteRow(AGridRow: Integer); reintroduce;
    procedure DragDrop(Source: TObject; X, Y: Integer); override;
    procedure EditingDone; override;
    function  EditorByStyle(Style: TColumnButtonStyle): TWinControl; override;
    procedure EndUpdate(ARefresh: Boolean = true);
    function GetGridCol(ASheetCol: Cardinal): Integer; inline;
    function GetGridRow(ASheetRow: Cardinal): Integer; inline;
    procedure GetSheets(const ASheets: TStrings);
    function GetWorksheetCol(AGridCol: Integer): Cardinal; inline;
    function GetWorksheetRow(AGridRow: Integer): Cardinal; inline;
    procedure InsertCol(AGridCol: Integer);
    procedure InsertRow(AGridRow: Integer);
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer = -1); overload;
    procedure LoadFromSpreadsheetFile(AFileName: string;
      AFormatID: TsSpreadFormatID = sfidUnknown; AWorksheetIndex: Integer = -1); overload;
    procedure LoadSheetFromSpreadsheetFile(AFileName: String;
      AWorksheetIndex: Integer = -1; AFormatID: TsSpreadFormatID = sfidUnknown);
    procedure LoadFromWorkbook(AWorkbook: TsWorkbook; AWorksheetIndex: Integer = -1);
    procedure NewWorkbook(AColCount, ARowCount: Integer);
    procedure SaveToSpreadsheetFile(AFileName: string;
      AOverwriteExisting: Boolean = true); overload;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormat: TsSpreadsheetFormat;
      AOverwriteExisting: Boolean = true); overload; deprecated;
    procedure SaveToSpreadsheetFile(AFileName: string; AFormatID: TsSpreadFormatID;
      AOverwriteExisting: Boolean = true); overload;
    procedure SelectSheetByIndex(AIndex: Integer);

    procedure MergeCells; overload;
    procedure MergeCells(ARect: TGridRect); overload;
    procedure MergeCells(ALeft, ATop, ARight, ABottom: Integer); overload;
    procedure UnmergeCells; overload;
    procedure UnmergeCells(ACol, ARow: Integer); overload;

    procedure ShowCellBorders(ALeft, ATop, ARight, ABottom: Integer;
      const ALeftOuterStyle, ATopOuterStyle, ARightOuterStyle, ABottomOuterStyle,
      AHorInnerStyle, AVertInnerStyle: TsCellBorderStyle);

    { Row height / col width calculation }
    procedure UpdateColWidth(ACol: Integer);
    procedure UpdateColWidths(AStartIndex: Integer = 0);
    procedure UpdateRowHeight(ARow: Integer; AEnforceCalcRowHeight: Boolean = false);
    procedure UpdateRowHeights(AStartRow: Integer = -1; AEnforceCalcRowHeight: Boolean = false);

    { Utilities related to Workbooks }
    procedure Convert_sFont_to_Font(sFont: TsFont; AFont: TFont);
    procedure Convert_Font_to_sFont(AFont: TFont; sFont: TsFont);

    { Interfacing with WorkbookSource}
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;

    { public properties }
    {@@ Link to the workbook }
    property WorkbookSource: TsWorkbookSource read GetWorkbookSource write SetWorkbookSource;
    {@@ Currently selected worksheet of the workbook }
    property Worksheet: TsWorksheet read GetWorksheet;
    {@@ Workbook displayed in the grid }
    property Workbook: TsWorkbook read GetWorkbook;
    {@@ Count of header lines - for conversion between grid- and workbook-based
     row and column indexes. Either 1 if row and column headers are shown or 0 if not}
    property HeaderCount: Integer read FHeaderCount;

    {@@ Background color of the cell at the given column and row. Expressed as
        index into the workbook's color palette. }
    property BackgroundColor[ACol, ARow: Integer]: TsColor
        read GetBackgroundColor write SetBackgroundColor;
    {@@ Common background color of the cells covered by the given rectangle.
        Expressed as index into the workbook's color palette. }
    property BackgroundColors[ALeft, ATop, ARight, ABottom: Integer]: TsColor
        read GetBackgroundColors write SetBackgroundColors;
    {@@ Override system or sheet right-to-left mode for cell }
    property CellBiDiMode[ACol,ARow: Integer]: TsBiDiMode
        read GetCellBiDiMode write SetCellBiDiMode;
    {@@ Set of flags indicating at which cell border a border line is drawn. }
    property CellBorder[ACol, ARow: Integer]: TsCellBorders
        read GetCellBorder write SetCellBorder;
    {@@ Set of flags indicating at which border of a range of cells a border
        line is drawn }
    property CellBorders[ALeft, ATop, ARight, ABottom: Integer]: TsCellBorders
        read GetCellBorders write SetCellBorders;
    {@@ Style of the border line at the given border of the cell at column ACol
        and row ARow. Requires the cellborder flag of the border to be set
        for the border line to be shown }
    property CellBorderStyle[ACol, ARow: Integer; ABorder: TsCellBorder]: TsCellBorderStyle
        read GetCellBorderStyle write SetCellBorderStyle;
    {@@ Style of the border line at the given border of the cells within the
        range of colum/row indexes defined by the rectangle. Requires the cellborder
        flag of the border to be set for the border line to be shown }
    property CellBorderStyles[ALeft, ATop, ARight, ABottom: Integer;
        ABorder: TsCellBorder]: TsCellBorderStyle
        read GetCellBorderStyles write SetCellBorderStyles;
    {@@ Comment assigned to the cell at column ACol and row ARow }
    property CellComment[ACol, ARow: Integer]: String
        read GetCellComment write SetCellComment;
    {@@ Font to be used for text in the cell at column ACol and row ARow. }
    property CellFont[ACol, ARow: Integer]: TFont
        read GetCellFont write SetCellFont;
    {@@ Font to be used for the cells in the column/row index range
        given by the rectangle }
    property CellFonts[ALeft, ATop, ARight, ABottom: Integer]: TFont
        read GetCellFonts write SetCellFonts;
    {@@ Color of the font used for the cell in column ACol and row ARow }
    property CellFontColor[ACol, ARow: Integer]: TsColor
        read GetCellFontColor write SetCellFontColor;
    {@@ Color of the font used for the cells within the range
        of column/row indexes defined by the rectangle, scUndefined if not constant. }
    property CellFontColors[ALeft, ATop, ARight, ABottom: Integer]: TsColor
        read GetCellFontColors write SetCellFontColors;
    {@@ Name of the font used for the cell in column ACol and row ARow }
    property CellFontName[ACol, ARow: Integer]: String
        read GetCellFontName write SetCellFontName;
    {@@ Name of the font used for the cells within the range
        of column/row indexes defined by the rectangle. }
    property CellFontNames[ALeft, ATop, ARight, ABottom: Integer]: String
        read GetCellFontNames write SetCellFontNames;
    {@@ Style of the font (bold, italic, ...) used for text in the
        cell at column ACol and row ARow. }
    property CellFontStyle[ACol, ARow: Integer]: TsFontStyles
        read GetCellFontStyle write SetCellFontStyle;
    {@@ Style of the font (bold, italic, ...) used for the cells within
        the range of column/row indexes defined by the rectangle. }
    property CellFontStyles[ALeft, ATop, ARight, ABottom: Integer]: TsFontStyles
        read GetCellFontStyles write SetCellFontStyles;
    {@@ Size of the font (in points) used for the cell at column ACol
        and row ARow }
    property CellFontSize[ACol, ARow: Integer]: Single
        read GetCellFontSize write SetCellFontSize;
    {@@ Size of the font (in points) used for the cells within the
        range of column/row indexes defined by the rectangle. }
    property CellFontSizes[ALeft, ATop, ARight, ABottom: Integer]: Single
        read GetCellFontSizes write SetCellFontSizes;
    {@@ Cell values }
    property Cells[ACol, ARow: Integer]: Variant
        read GetCellValue write SetCellValue;
    {@@ Parameter for horizontal text alignment within the cell at column ACol
        and row ARow }
    property HorAlignment[ACol, ARow: Integer]: TsHorAlignment
        read GetHorAlignment write SetHorAlignment;
    {@@ Parameter for the horizontal text alignments in all cells within the
        range cf column/row indexes defined by the rectangle. }
    property HorAlignments[ALeft, ATop, ARight, ABottom: Integer]: TsHorAlignment
        read GetHorAlignments write SetHorAlignments;
    {@@ Hyperlink assigned to the cell in row ARow and column ACol }
    property Hyperlink[ACol, ARow: Integer]: String
        read GetHyperlink write SetHyperlink;
    {@@ Number format (as Excel string) to be applied to cell at column ACol and row ARow. }
    property NumberFormat[ACol, ARow: Integer]: String
        read GetNumberFormat write SetNumberFormat;
    {@@ Number format (as Excel string) to be applied to all cells within the
        range of column/row indexes defined by the rectangle. }
    property NumberFormats[ALeft, ATop, ARight, ABottom: Integer]: String
        read GetNumberFormats write SetNumberFormats;
    {@@ Rotation of the text in the cell at column ACol and row ARow. }
    property TextRotation[ACol, ARow: Integer]: TsTextRotation
        read GetTextRotation write SetTextRotation;
    {@@ Rotation of the text in the cells within the range of column/row indexes
        defined by the rectangle. }
    property TextRotations[ALeft, ATop, ARight, ABottom: Integer]: TsTextRotation
        read GetTextRotations write SetTextRotations;
    {@@ Parameter for vertical text alignment in the cell at column ACol and
        row ARow. }
    property VertAlignment[ACol, ARow: Integer]: TsVertAlignment
        read GetVertAlignment write SetVertAlignment;
    {@@ Parameter for vertical text alignment in the cells having column/row
        indexes defined by the rectangle. }
    property VertAlignments[ALeft, ATop, ARight, ABottom: Integer]: TsVertAlignment
        read GetVertAlignments write SetVertAlignments;
    {@@ If true, word-wrapping of text within the cell at column ACol and row ARow
        is activated. }
    property Wordwrap[ACol, ARow: Integer]: Boolean
        read GetWordwrap write SetWordwrap;
    {@@ If true, word-wrapping of text within all cells within the range defined
        by the rectangle is activated. }
    property Wordwraps[ALeft, ATop, ARight, ABottom: Integer]: Boolean
        read GetWordwraps write SetWordwraps;
    {@@ Zoomfactor of the grid }
    property ZoomFactor: Double
        read GetZoomFactor write SetZoomFactor;

    // inherited, but modified

    {@@ Column width, in pixels }
    property ColWidths[ACol: Integer]: Integer
        read GetColWidths write SetColWidths;
    {@@ Default column width, in pixels }
    property DefaultColWidth: Integer
        read GetDefColWidth write SetDefColWidth;
    {@@ Default row height, in pixels }
    property DefaultRowHeight: Integer
        read GetDefRowHeight write SetDefRowHeight;
    {@@ Row height in pixels }
    property RowHeights[ARow: Integer]: Integer
        read GetRowHeights write SetRowHeights;

    // inherited

   {$IFNDEF FPS_NO_GRID_MULTISELECT}
    {@@ Allow multiple selections}
    property RangeSelectMode default rsmMulti;
   {$ENDIF}
  end;


  { TsWorksheetGrid }

  {@@
    TsWorksheetGrid is a grid which displays spreadsheet data along with
    formatting. As it is linked to an instance of TsWorkbook, it provides
    methods for reading data from or writing to spreadsheet files. It has the
    same funtionality as TsCustomWorksheetGrid, but has published all properties.
  }
  TsWorksheetGrid = class(TsCustomWorksheetGrid)
  published
    // inherited from TsCustomWorksheetGrid
    {@@ Allow built-in drag and drop }
    property AllowDragAndDrop;
    {@@ Automatically recalculates the worksheet formulas if a cell value changes. }
    property AutoCalc;
    {@@ Automatically expand grid dimensions }
    property AutoExpand;
    {@@ Displays column and row headers in the fixed col/row style of the grid.
        Deprecated. Use ShowHeaders instead. }
    property DisplayFixedColRow; deprecated 'Use ShowHeaders';
    {@@ Determines whether a single- or multiline cell editor is used }
    property EditorLineMode;
    {@@ Pen for the line separating the frozen cols/rows from the regular grid }
    property FrozenBorderPen;
    {@@ This number of columns at the left is "frozen", i.e. it is not possible to
        scroll these columns. }
    property FrozenCols;
    {@@ This number of rows at the top is "frozen", i.e. it is not possible to
        scroll these rows. }
    property FrozenRows;
    {@@ Activates reading of RPN formulas. Should be turned off when
        non-implemented formulas crashe reading of the spreadsheet file. }
    property ReadFormulas;
    {@@ Pen used for drawing the selection rectangle }
    property SelectionPen;
    {@@ Shows/hides vertical and horizontal grid lines. }
    property ShowGridLines;
    {@@ Shows/hides column and row headers in the fixed col/row style of the grid. }
    property ShowHeaders;
    {@@ Activates text overflow (cells reaching into neighbors) }
    property TextOverflow;
    {@@ Link to the workbook }
    property WorkbookSource;

    {@@ inherited from ancestors}
    property Align;
    {@@ inherited from ancestors}
    property AlternateColor;
    {@@ inherited from ancestors}
    property Anchors;
    {@@ inherited from ancestors}
    property AutoAdvance;
    {@@ inherited from ancestors}
    property AutoEdit;
    {@@ inherited from ancestors}
    property AutoFillColumns;
    //property BiDiMode;
    {@@ inherited from ancestors}
    property BorderSpacing;
    {@@ inherited from ancestors}
    property BorderStyle;
    {@@ inherited from ancestors}
    property CellHintPriority;
    {@@ inherited from ancestors}
    property Color;
    {@@ inherited from ancestors}
    property ColCount default 27; //stored;
    //property Columns;
    {@@ inherited from ancestors}
    property Constraints;
    {@@ inherited from ancestors}
    property DefaultColWidth;
    {@@ inherited from ancestors}
    property DefaultDrawing;
    {@@ inherited from ancestors}
    property DefaultRowHeight;
    {@@ inherited from ancestors}
    property DragCursor;
    {@@ inherited from ancestors}
    property DragKind;
    {@@ inherited from ancestors}
    property DragMode;
    {@@ inherited from ancestors}
    property Enabled;
    {@@ inherited from ancestors}
    property ExtendedSelect default true;
    {@@ inherited from ancestors}
    property FixedColor;
    {@@ inherited from ancestors}
    property Flat;
    {@@ inherited from ancestors}
    property Font;
    {@@ inherited from ancestors}
    property GridLineWidth;
    {@@ inherited from ancestors}
    property HeaderHotZones;
    {@@ inherited from ancestors}
    property HeaderPushZones;
    {@@ inherited from ancestors}
    property MouseWheelOption;
    {@@ inherited from TCustomGrid. Select the option goEditing to make the grid editable! }
    property Options;
    {@@ inherited from ancestors }
    property ParentBiDiMode;
    {@@ inherited from ancestors}
    property ParentColor default false;
    {@@ inherited from ancestors}
    property ParentFont;
    {@@ inherited from ancestors}
    property ParentShowHint;
    {@@ inherited from ancestors}
    property PopupMenu;
    {@@ inherited from ancestors}
    property RowCount default 101;
    {@@ inherited from ancestors}
    property ScrollBars;
    {@@ inherited from ancestors}
    property ShowHint;
    {@@ inherited from ancestors}
    property TabOrder;
    {@@ inherited from ancestors}
    property TabStop;
    {@@ inherited from ancestors}
    property TitleFont;
    {@@ inherited from ancestors}
    property TitleImageList;
    {@@ inherited from ancestors}
    property TitleStyle;
    {@@ inherited from ancestors}
    property UseXORFeatures;
    {@@ inherited from ancestors}
    property Visible;
    {@@ inherited from ancestors}
    property VisibleColCount;
    {@@ inherited from ancestors}
    property VisibleRowCount;

    {@@ inherited from ancestors}
    property OnBeforeSelection;
    {@@ inherited from ancestors}
    property OnChangeBounds;
    {@@ inherited from ancestors}
    property OnClick;
    {@@ inherited from TCustomWorksheetGrid}
    property OnClickHyperlink;
    {@@ inherited from ancestors}
    property OnColRowDeleted;
    {@@ inherited from ancestors}
    property OnColRowExchanged;
    {@@ inherited from ancestors}
    property OnColRowInserted;
    {@@ inherited from ancestors}
    property OnColRowMoved;
    (*
    {@@ inherited from ancestors}
    property OnCompareCells;     // apply userdefined sorting to worksheet directly!
    *)
    {@@ inherited from ancestors}
    property OnDragDrop;
    {@@ inherited from ancestors}
    property OnDragOver;
    {@@ inherited from ancestors}
    property OnDblClick;
    {@@ inherited from ancestors}
    property OnDrawCell;
    {@@ inherited from ancestors}
    property OnEditButtonClick;
    {@@ inherited from ancestors}
    property OnEditingDone;
    {@@ inherited from ancestors}
    property OnEndDock;
    {@@ inherited from ancestors}
    property OnEndDrag;
    {@@ inherited from ancestors}
    property OnEnter;
    {@@ inherited from ancestors}
    property OnExit;
    {@@ inherited from ancestors}
    property OnGetEditMask;
    {@@ inherited from ancestors}
    property OnGetEditText;
    {@@ inherited from ancestors}
    property OnHeaderClick;
    {@@ inherited from ancestors}
    property OnHeaderSized;
    {@@ inherited from ancestors}
    property OnHeaderSizing;
    {@@ inherited from ancestors}
    property OnKeyDown;
    {@@ inherited from ancestors}
    property OnKeyPress;
    {@@ inherited from ancestors}
    property OnKeyUp;
    {@@ inherited from ancestors}
    property OnMouseDown;
    {@@ inherited from ancestors}
    property OnMouseMove;
    {@@ inherited from ancestors}
    property OnMouseUp;
    {@@ inherited from ancestors}
    property OnMouseWheel;
    {@@ inherited from ancestors}
    property OnMouseWheelDown;
    {@@ inherited from ancestors}
    property OnMouseWheelUp;
    {@@ inherited from ancestors}
    property OnPickListSelect;
    {@@ inherited from ancestors}
    property OnPrepareCanvas;
    {@@ inherited from ancestors}
    property OnResize;
    {@@ inherited from ancestors}
    property OnSelectEditor;
    {@@ inherited from ancestors}
    property OnSelection;
    {@@ inherited from ancestors}
    property OnSelectCell;
    {@@ inherited from ancestors}
    property OnSetEditText;
    {@@ inherited from ancestors}
    property OnShowHint;
    {@@ inherited from ancestors}
    property OnStartDock;
    {@@ inherited from ancestors}
    property OnStartDrag;
    {@@ inherited from ancestors}
    property OnTopLeftChanged;
    {@@ inherited from ancestors}
    property OnUTF8KeyPress;
    {@@ inherited from ancestors}
    property OnValidateEntry;
    {@@ inherited from ancestors}
    property OnContextPopup;
  end;

const
  NO_CELL_BORDER: TsCellBorderStyle = (LineStyle: lsThin; Color: scNotDefined);

var
  {@@ Default number of columns prepared for a new empty worksheet }
  DEFAULT_COL_COUNT: Integer = 26;

  {@@ Default number of rows prepared for a new empty worksheet }
  DEFAULT_ROW_COUNT: Integer = 100;

  {@@ Tolerance for mouse for hitting cell border }
  CELL_BORDER_DELTA: Integer = 4;

  {@@ Cursor for copy operation during drag and drop }
  crDragCopy: Integer;

procedure Register;


implementation

uses
  Types, LCLType, LCLIntf, LCLProc, LazUTF8, Math, StrUtils,
  fpCanvas, {%H-}fpsPatches,
  fpsStrings, fpsUtils, fpsVisualUtils, fpsHTMLUtils, fpsImages, fpsNumFormat;

const
  {@@ Interval how long the mouse buttons has to be held down on a
    hyperlink cell until the associated hyperlink is executed. }
  HYPERLINK_TIMER_INTERVAL = 500;

const
  // Constants for AGridPart parameter
  DRAW_NON_FROZEN = 0;
  DRAW_FROZEN_ROWS = 1;
  DRAW_FROZEN_COLS = 2;
  DRAW_FROZEN_CORNER = 3;

var
  {@@ Auxiliary bitmap containing the previously used non-trivial fill pattern }
  FillPatternBitmap: TBitmap = nil;
  FillPatternStyle: TsFillStyle;
  FillPatternFgColor: TColor;
  FillPatternBgColor: TColor;

  DragBorderBitmap: TBitmap = nil;

{@@ ----------------------------------------------------------------------------
  Helper procedure which creates bitmaps used for fill patterns in cell
  backgrounds.
  The parameters are buffered in FillPatternXXXX variables to avoid unnecessary
  creation of the same bitmaps again and again.
-------------------------------------------------------------------------------}
procedure CreateFillPattern(var ABitmap: TBitmap; AStyle: TsFillStyle;
  AFgColor, ABgColor: TColor);

  procedure SolidFill(AColor: TColor);
  begin
    ABitmap.Canvas.Brush.Color := AColor;
    ABitmap.Canvas.FillRect(0, 0, ABitmap.Width, ABitmap.Height);
  end;

var
  x,y: Integer;
begin
  if (FillPatternStyle = AStyle) and (FillPatternBgColor = ABgColor) and
     (FillPatternFgColor = AFgColor) and (ABitmap <> nil)
  then
    exit;

  FreeAndNil(ABitmap);
  ABitmap := TBitmap.Create;
  with ABitmap do begin
    if AStyle = fsGray6 then SetSize(8, 4) else SetSize(4, 4);
    case AStyle of
      fsNoFill:
        SolidFill(ABgColor);
      fsSolidFill:
        SolidFill(AFgColor);
      fsGray75:
        begin
          SolidFill(AFgColor);
          Canvas.Pixels[0, 0] := ABgColor;
          Canvas.Pixels[2, 1] := ABgColor;
          Canvas.Pixels[0, 2] := ABgColor;
          Canvas.Pixels[2, 3] := ABgColor;
        end;
      fsGray50:
        begin
          SolidFill(AFgColor);
          for y := 0 to 3 do for
            x := 0 to 3 do
              if odd(x+y) then Canvas.Pixels[x,y] := ABgColor;
        end;
      fsGray25:
        begin
          SolidFill(ABgColor);
          Canvas.Pixels[0, 0] := AFgColor;
          Canvas.Pixels[2, 1] := AFgColor;
          Canvas.Pixels[0, 2] := AFgColor;
          Canvas.Pixels[2, 3] := AFgColor;
        end;
      fsGray12:
        begin
          SolidFill(ABgColor);
          Canvas.Pixels[0, 0] := AFgColor;
          Canvas.Pixels[2, 2] := AFgColor;
        end;
      fsGray6:
        begin
          SolidFill(ABgColor);
          Canvas.Pixels[0, 0] := AFgColor;
          Canvas.Pixels[4, 2] := AFgColor;
        end;
      fsStripeHor:
        begin
          SolidFill(ABgColor);
          for y := 0 to 1 do
            for x := 0 to 3 do
              Canvas.Pixels[x,y] := AFgColor;
        end;
      fsStripeVert:
        begin
          SolidFill(ABgColor);
          for y := 0 to 3 do
            for x := 0 to 1 do
              Canvas.Pixels[x,y] := AFgColor;
        end;
      fsStripeDiagUp:
        begin
          SolidFill(ABgColor);
          for y := 0 to 3 do
            for x := 0 to 1 do
              Canvas.Pixels[(x+y) mod 4, 3-y] := AFgColor;
        end;
      fsStripeDiagDown:
        begin
          SolidFill(ABgColor);
          for y := 0 to 3 do
            for x := 0 to 1 do
              Canvas.Pixels[(x+y) mod 4, y] := AFgColor;
        end;
      fsThinStripeHor:
        begin
          SolidFill(ABgColor);
          for x := 0 to 3 do Canvas.Pixels[x, 0] := AFgColor;
        end;
      fsThinStripeVert:
        begin
          SolidFill(ABgColor);
          for y := 0 to 3 do Canvas.Pixels[0, y] := AFgColor;
        end;
      fsThinStripeDiagUp:
        begin
          SolidFill(ABgColor);
          for x := 0 to 3 do Canvas.Pixels[3-x, x] := AFgColor;
        end;
      fsThinStripeDiagDown, fsThinHatchDiag:
        begin
          SolidFill(ABgColor);
          for x := 0 to 3 do Canvas.Pixels[x, x] := AFgColor;
          if AStyle = fsThinHatchDiag then begin
            Canvas.Pixels[0, 2] := AFgColor;
            Canvas.Pixels[2, 0] := AFgColor;
          end;
        end;
      fsHatchDiag:
        begin
          SolidFill(ABgColor);
          for x := 0 to 1 do
            for y := 0 to 1 do begin
              Canvas.Pixels[x,y] := AFgColor;
              Canvas.Pixels[x+2, y+2] := AFgColor;
            end;
        end;
      fsThickHatchDiag:
        begin
          SolidFill(AFgColor);
          for x := 2 to 3 do Canvas.Pixels[x, 0] := ABgColor;
          for x := 0 to 1 do Canvas.Pixels[x, 2] := ABgColor;
        end;
      fsThinHatchHor:
        begin
          SolidFill(ABgColor);
          for x := 0 to 3 do begin
            Canvas.Pixels[x, 0] := AFgColor;
            Canvas.Pixels[0, x] := AFgColor;
          end;
        end;
    end;  // case
  end;

  FillPatternStyle := AStyle;
  FillPatternBgColor := ABgColor;
  FillPatternFgColor := AFgColor;
end;

                                    (*
{@@ ----------------------------------------------------------------------------
  Helper procedure which draws a densely dotted horizontal line. In Excel
  this is called a "hair line".

  @param x1, x2   x coordinates of the end points of the line
  @param y        y coordinate of the horizontal line
-------------------------------------------------------------------------------}
procedure DrawHairLineHor(ACanvas: TCanvas; x1, x2, y: Integer);
var
  clr: TColor;
  x: Integer;
begin
  if odd(x1) then inc(x1);
  x := x1;
  clr := ACanvas.Pen.Color;
  while (x <= x2) do begin
    ACanvas.Pixels[x, y] := clr;
    inc(x, 2);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Helper procedure which draws a densely dotted vertical line. In Excel
  this is called a "hair line".

  @param x        x coordinate of the vertical line
  @param y1, y2   y coordinates of the end points of the line
-------------------------------------------------------------------------------}
procedure DrawHairLineVert(ACanvas: TCanvas; x, y1, y2: Integer);
var
  clr: TColor;
  y: Integer;
begin
  if odd(y1) then inc(y1);
  y := y1;
  clr := ACanvas.Pen.Color;
  while (y <= y2) do begin
    ACanvas.Pixels[x, y] := clr;
    inc(y, 2);
  end;
end;                                  *)

{@@ ----------------------------------------------------------------------------
  Calculates a background color for selected cells. The procedures takes the
  original background color and dims or brightens it by adding the value ADelta
  to the RGB components.

  @param  c       Color to be modified
  @param  ADelta  Value to be added to the RGB components of the inpur color
  @result Modified color.
-------------------------------------------------------------------------------}
function CalcSelectionColor(c: TColor; ADelta: Byte) : TColor;
type
  TRGBA = record R,G,B,A: Byte end;
begin
  c := ColorToRGB(c);
  TRGBA(Result).A := 0;
  if TRGBA(c).R < 128
    then TRGBA(Result).R := TRGBA(c).R + ADelta
    else TRGBA(Result).R := TRGBA(c).R - ADelta;
  if TRGBA(c).G < 128
    then TRGBA(Result).G := TRGBA(c).G + ADelta
    else TRGBA(Result).G := TRGBA(c).G - ADelta;
  if TRGBA(c).B < 128
    then TRGBA(Result).B := TRGBA(c).B + ADelta
    else TRGBA(Result).B := TRGBA(c).B - ADelta;
end;

function VerticalIntersect(const ARect, BRect: TRect): Boolean;
begin
  Result := (ARect.Top < BRect.Bottom) and (ARect.Bottom > BRect.Top);
end;

function HorizontalIntersect(const ARect, BRect: TRect): Boolean;
begin
  Result := (ARect.Left < BRect.Right) and (ARect.Right > BRect.Left);
end;


{*******************************************************************************
*                                   TsSelPen                                   *
*******************************************************************************}
constructor TsSelPen.Create;
begin
  inherited;
  Width := 3;
  JoinStyle := pjsMiter;
end;


{*******************************************************************************
*                            TMultilineStringCellEditor                        *
*******************************************************************************}

constructor TMultilineStringCellEditor.Create(Aowner: TComponent);
begin
  inherited Create(AOwner);
  AutoSize := false;
end;

procedure TMultilineStringCellEditor.Change;
begin
  inherited Change;
  if (FGrid <> nil) and Visible then
    FGrid.EditorTextChanged(FCol, FRow, Text);
end;

procedure TMultilineStringCellEditor.EditingDone;
begin
  inherited EditingDone;
  if FGrid <> nil then
    FGrid.EditingDone;
end;

procedure TMultilineStringCellEditor.KeyDown(var Key: Word; Shift: TShiftState);

  function AllSelected: boolean;
  begin
    result := (SelLength > 0) and (SelLength = UTF8Length(Text));
  end;

  function AtStart: Boolean;
  begin
    Result:= (SelStart = 0);
  end;

  function AtEnd: Boolean;
  begin
    result := ((SelStart + 1) > UTF8Length(Text)) or AllSelected;
  end;

  procedure doEditorKeyDown;
  begin
    if FGrid <> nil then
      FGrid.EditorkeyDown(Self, key, shift);
  end;

  procedure doGridKeyDown;
  begin
    if FGrid <> nil then
      FGrid.KeyDown(Key, shift);
  end;

  function GetFastEntry: boolean;
  begin
    if FGrid <> nil then
      Result := FGrid.FastEditing
    else
      Result := False;
  end;

  procedure CheckEditingKey;
  begin
    if (FGrid = nil) or FGrid.EditorIsReadOnly then
      Key := 0;
  end;

var
  IntSel: boolean;
begin
  inherited KeyDown(Key,Shift);
  case Key of
    VK_F2:
      if AllSelected then begin
        SelLength := 0;
        SelStart := Length(Text);
      end;
    VK_DELETE, VK_BACK:
      CheckEditingKey;
    VK_UP, VK_DOWN:
      doGridKeyDown;
    VK_LEFT, VK_RIGHT:
      if GetFastEntry then begin
        IntSel:=
          ((Key=VK_LEFT) and not AtStart) or
          ((Key=VK_RIGHT) and not AtEnd);
      if not IntSel then begin
          doGridKeyDown;
      end;
    end;
    VK_END, VK_HOME:
      ;
    VK_ESCAPE:
      begin
        doGridKeyDown;
        if key<>0 then begin
          Text := FGrid.FOldEditorText;
//          Lines.Text := '';   // FIXME: FGrid.FEditorOldvalue;
//          SetEditText(FGrid.FEditorOldValue);
          FGrid.EditorHide;
        end;
      end;
    else
      doEditorKeyDown;
  end;
end;

procedure TMultilineStringCellEditor.msg_GetGrid(var Msg: TGridMessage);
begin
  Msg.Grid := FGrid;
  Msg.Options:= EO_IMPLEMENTED;
end;

procedure TMultilineStringCellEditor.msg_GetValue(var Msg: TGridMessage);
begin
  Msg.Col := FCol;
  Msg.Row := FRow;
  Msg.Value := Text;
end;

procedure TMultilineStringCellEditor.msg_SelectAll(var Msg: TGridMessage);
begin
  Unused(Msg);
  SelectAll;
end;

procedure TMultilineStringCellEditor.msg_SetGrid(var Msg: TGridMessage);
begin
  FGrid := Msg.Grid as TsCustomWorksheetGrid;
  Msg.Options := EO_AUTOSIZE or EO_SELECTALL or EO_HOOKKEYPRESS or EO_HOOKKEYUP;
end;

procedure TMultilineStringCellEditor.msg_SetPos(var Msg: TGridMessage);
begin
  FCol := Msg.Col;
  FRow := Msg.Row;
end;

procedure TMultilineStringCellEditor.msg_SetValue(var Msg: TGridMessage);
begin
  Text := Msg.Value;
  SelStart := UTF8Length(Text);
end;

procedure TMultilineStringCellEditor.WndProc(var TheMessage: TLMessage);
begin
  if FGrid <> nil then
    case TheMessage.Msg of
      LM_CLEAR, LM_CUT, LM_PASTE:
        if FGrid.EditorIsReadOnly then exit;
    end;
  inherited WndProc(TheMessage);
end;


{*******************************************************************************
*                              TsCustomWorksheetGrid                           *
*******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Constructor of the grid. Activates the display of column and row headers
  and creates an internal "CellFont". Creates a pre-defined number of empty rows
  and columns.

  @param  AOwner   Owner of the grid
-------------------------------------------------------------------------------}
constructor TsCustomWorksheetGrid.Create(AOwner: TComponent);
begin
  inc(FRowHeightLock);

  FInternalWorkbookSource := TsWorkbookSource.Create(self);
  FInternalWorkbookSource.Name := 'internal';

  inc(FActiveCellLock);
  inherited Create(AOwner);
  dec(FActiveCellLock);

  AutoAdvance := aaDown;
  ExtendedSelect := true;
  FHeaderCount := 1;
  ColCount := DEFAULT_COL_COUNT + FHeaderCount;
  RowCount := DEFAULT_ROW_COUNT + FHeaderCount;

  FDefRowHeight100 := inherited GetDefaultRowHeight;
  FDefColWidth100 := inherited DefaultColWidth;

  FCellFont := TFont.Create;
  FSelPen := TsSelPen.Create;
  FSelPen.Style := psSolid;
  FSelPen.Color := clBlack;
  FSelPen.JoinStyle := pjsMiter;
  FSelPen.OnChange := @GenericPenChangeHandler;

  FFrozenBorderPen := TPen.Create;
  FFrozenBorderPen.Style := psSolid;
  FFrozenBorderPen.Color := clBlack;
  FFrozenBorderPen.OnChange := @GenericPenChangeHandler;

  FAutoExpand := [aeData, aeNavigation, aeDefault];
  FHyperlinkTimer := TTimer.Create(self);
  FHyperlinkTimer.Interval := HYPERLINK_TIMER_INTERVAL;
  FHyperlinkTimer.OnTimer := @HyperlinkTimerElapsed;
  SetWorkbookSource(FInternalWorkbookSource);
 {$IFNDEF FPS_NO_GRID_MULTISELECT}
  RangeSelectMode := rsmMulti;
 {$ENDIF}

  FAllowDragAndDrop := true;

  dec(FRowHeightLock);
  UpdateRowHeights;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the grid: Destroys the workbook and the internal CellFont.
-------------------------------------------------------------------------------}
destructor TsCustomWorksheetGrid.Destroy;
begin
  SetWorkbookSource(nil);
  if FInternalWorkbookSource <> nil then
    FInternalWorkbookSource.RemoveListener(self);  // will be destroyed automatically
  FreeAndNil(FCellFont);
  FreeAndNil(FSelPen);
  FreeAndNil(FMultilineStringEditor);
  inherited Destroy;
end;

procedure TsCustomWorksheetGrid.AdaptToZoomFactor;
var
  c, r: Integer;
begin
  inc(FZoomLock);
  DefaultRowHeight := round(GetZoomfactor * FDefRowHeight100);
  DefaultColWidth := round(GetZoomFactor * FDefColWidth100);
  UpdateColWidths;
  UpdateRowHeights;
  dec(FZoomLock);

  // Bring active cell back into the viewport: There is a ScrollToCell but
  // this method is private. It is called by SetCol/SetRow, though.
  if ((Col < GCache.Visiblegrid.Left) or (Col >= GCache.VisibleGrid.Right)) and
     (GCache.VisibleGrid.Left <> GCache.VisibleGrid.Right) then
  begin
    c := Col;
    Col := c-1;    // "Col" must change in order to call ScrtollToCell
    Col := c;
  end;
  if ((Row < GCache.VisibleGrid.Top) or (Row >= GCache.VisibleGrid.Bottom)) and
     (GCache.VisibleGrid.Top <> GCache.VisibleGrid.Bottom) then
  begin
    r := Row;
    Row := r-1;
    Row := r;
  end;
end;

procedure TsCustomWorksheetGrid.AutoColWidth(ACol: Integer);
begin
  AutoAdjustColumn(ACol);
end;

procedure TscustomWorksheetGrid.AutoRowHeight(ARow: Integer);
begin
  AutoAdjustRow(ARow);
end;

{@@ ----------------------------------------------------------------------------
  Is called when goDblClickAutoSize is in the grid's options and a double click
  has occured at the border of a column header. Sets optimum column with.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.AutoAdjustColumn(ACol: Integer);
var
  gRow: Integer;  // row in grid coordinates
  w, maxw: Integer;
  txt: String;
  cell: PCell;
  RTL: Boolean;
begin
  if Worksheet = nil then
    exit;

  RTL := IsRightToLeft;
  maxw := -1;
  for cell in Worksheet.Cells.GetColEnumerator(GetWorkSheetCol(ACol)) do
  begin
    // Merged cells are not considered for calculating AutoColWidth -- see Excel.
    if Worksheet.IsMerged(cell) then
      continue;
    gRow := GetGridRow(cell^.Row);
    txt := GetCellText(ACol, gRow, false);
    if txt = '' then
      Continue;
    case Worksheet.ReadBiDiMode(cell) of
      bdRTL: RTL := true;
      bdLTR: RTL := false;
    end;
    w := RichTextWidth(Canvas, Workbook, Rect(0, 0, MaxInt, MaxInt),
      txt, cell^.RichTextParams, Worksheet.ReadCellFontIndex(cell),
      Worksheet.ReadTextRotation(cell), false, RTL, ZoomFactor);
    if w > maxw then maxw := w;
  end;
  if maxw > -1 then
    maxw := maxw + 2*constCellPadding
  else
    maxw := DefaultColWidth;
  ColWidths[ACol] := maxW;
  HeaderSized(true, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Is called when goDblClickAutoSize is in the grid's options and a double click
  has occured at the border of a row header. Sets optimum row height.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.AutoAdjustRow(ARow: Integer);
begin
  inc(FZoomLock);
  if Worksheet <> nil then
    RowHeights[ARow] := CalcAutoRowHeight(ARow)
  else
    RowHeights[ARow] := DefaultRowHeight;
  HeaderSized(false, ARow);
  dec(FZoomLock);
end;

{@@ ----------------------------------------------------------------------------
  Automatically expands the ColCount such that the specified column fits in
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.AutoExpandToCol(ACol: Integer;
  AMode: TsAutoExpandMode);
begin
  if ACol >= ColCount then
  begin
    if (AMode in FAutoExpand) then
      ColCount := ACol + 1
    else
      raise Exception.CreateFmt(rsOperationExceedsColCount, [ACol, ColCount]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Automatically expands the RowCount such that the specified column fits in
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.AutoExpandToRow(ARow: Integer;
  AMode: TsAutoExpandMode);
begin
  if ARow >= RowCount then
  begin
    if (AMode in FAutoExpand) then
      RowCount := ARow + 1
    else
      raise Exception.CreateFmt(rsOperationExceedsRowCount, [ARow, RowCount]);
  end;
end;

{@@ ----------------------------------------------------------------------------
  The BeginUpdate/EndUpdate pair suppresses unnecessary painting of the grid.
  Call BeginUpdate to stop refreshing the grid, and call EndUpdate to release
  the lock and to repaint the grid again.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.BeginUpdate;
begin
  inc(FLockCount);
  inherited BeginUpdate;
end;

{@@ ----------------------------------------------------------------------------
  Converts the column width, given in units used by the worksheet, to pixels.

  @param   AWidth   Width of a column in units used by the worksheet
  @return  Column width in pixels.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcColWidthFromSheet(AWidth: Single): Integer;
var
  w_pts: Double;
begin
  w_pts := Workbook.ConvertUnits(AWidth, Workbook.Units, suPoints);
  Result := PtsToPx(w_pts, Screen.PixelsPerInch);
end;

{@@ ----------------------------------------------------------------------------
  Finds the maximum cell height per row and uses this to define the RowHeights[].
  Returns DefaultRowHeight if the row does not contain any cells, or if the
  worksheet does not have a TRow record for this particular row.
  ARow is a grid row index.

  @param   ARow  Index of the row, in grid units
  @return  Row height
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcAutoRowHeight(ARow: Integer): Integer;
var
  c: Integer;
  h: Integer;
begin
  h := 0;
  for c := FHeaderCount to ColCount-1 do
    h := Max(h, GetCellHeight(c, ARow));  // Zoom factor is applied to font size
  if h = 0 then
    Result := DefaultRowHeight            // Zoom factor applied by getter function
  else
    Result := h;
end;

{@@ ----------------------------------------------------------------------------
  Converts the row height (from a worksheet row record), given in units used by
  the sheet, to pixels as needed by the grid

  @param  AHeight  Row height expressed in units used by the worksheet.
  @result Row height in pixels.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcRowHeightFromSheet(AHeight: Single): Integer;
var
  h_pts: Single;
begin
  h_pts := Workbook.ConvertUnits(abs(AHeight), Workbook.Units, suPoints);;
  Result := PtsToPx(h_pts, Screen.PixelsPerInch); // + 4;
end;

function TsCustomWorksheetGrid.CalcRowHeightToSheet(AHeight: Integer): Single;
var
  h_pts: Single;
begin
  h_pts := PxToPts(AHeight, Screen.PixelsPerInch);
  Result := Workbook.ConvertUnits(h_pts, suPoints, Workbook.Units);
end;

{@@ ----------------------------------------------------------------------------
  Calculates the top-left corner (in pixels) of the area which can be
  scrolled. Is bordered by the fixed header cells and the frozen columns and
  rows.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcTopLeft(AHeaderOnly: Boolean): TPoint;
var
  fc, fr: Integer;
  tmp: Integer = 0;
begin
  Result := Point(0, 0);  // to silence the compiler...

  fc := IfThen(AHeaderOnly, FHeaderCount, FHeaderCount + FFrozenCols);
  if IsRightToLeft then
  begin
    if fc > 0 then
      ColRowToOffset(true, true, fc-1, Result.X, tmp)
    else
      Result.X := ClientRect.Right;
  end else
  begin
    if fc > 0 then
      ColRowToOffset(true, true, fc-1, tmp, Result.X)
    else
      Result.X := ClientRect.Left;
  end;

  fr := IfThen(AHeaderOnly, FHeaderCount, FHeaderCount + FFrozenRows);
  if fr > 0 then
    ColRowToOffset(false, true, fr-1, tmp, Result.Y)
  else
    Result.Y := ClientRect.Top;
end;

{@@ ----------------------------------------------------------------------------
  Converts the column height given in screen pixels to the units used by the
  worksheet.

  @param   AValue   Column width in pixels
  @result  Column width expressed in units defined by the workbook.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcWorksheetColWidth(AValue: Integer): Single;
var
  w_pts: Double;
begin
  Result := 0;
  if Worksheet <> nil then
  begin
    // The grid's column width is in "pixels", the worksheet's column width
    // has the units defined by the workbook.
    w_pts := PxToPts(AValue/ZoomFactor, Screen.PixelsPerInch);
    Result := Workbook.ConvertUnits(w_pts, suPoints, Workbook.Units);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Converts the row height given in screen pixels to the units used by the
  worksheet.

  @param   AValue   Row height in pixels
  @result  Row height expressed in units defined by the workbook.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CalcWorksheetRowHeight(AValue: Integer): Single;
var
  h_pts: Double;
begin
  Result := 0;
  if Worksheet <> nil then
  begin
    // The grid's row heights are in "pixels", the worksheet's row height
    // has the units defined by the workbook.
    h_pts := PxToPts(AValue/ZoomFactor, Screen.PixelsPerInch);
    Result := Workbook.ConvertUnits(h_pts, suPoints, Workbook.Units);
  end;
end;

function TsCustomWorksheetGrid.CanEditShow: Boolean;
begin
  Result := inherited and (not FReadOnly);
end;

{@@ ----------------------------------------------------------------------------
  Looks for overflowing cells: if the text of the given cell is longer than
  the cell width the function calculates the column indexes and the rectangle
  to show the complete text.
  Ony for non-wordwrapped label cells and for horizontal orientation.
  Function returns false if text overflow needs not to be considered.

  @param ACol, ARow   Column and row indexes (in grid coordinates) of the cell
                      to be drawn
  @param AState       GridDrawState of the cell (normal, fixed, selected etc)
  @param ACol1,ACol2  (output) Index of the first and last column covered by the
                      overflowing text
  @param ARect        (output) Pixel rectangle enclosing the cell and its neighbors
                      affected
  @return TRUE if text overflow into neighbor cells is to be considered,
          FALSE if not.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.CellOverflow(ACol, ARow: Integer;
  AState: TGridDrawState; out ACol1, ACol2: Integer; var ARect: TRect): Boolean;
var
  txt: String;
  len: Integer;
  cell: PCell;
  txtalign: TsHorAlignment;
  r: Cardinal;
  w, w0: Integer;
  fmt: PsCellFormat;
  fc: Integer;
begin
  Result := false;
  cell := FDrawingCell;

  // Nothing to do in these cases (like in Excel):
  if (cell = nil) or not (cell^.ContentType in [cctUTF8String]) then  // ... non-label cells
    exit;

//  fmt := Workbook.GetPointerToCellFormat(cell^.FormatIndex);
  fmt := Worksheet.GetPointerToEffectiveCellFormat(cell);
  if (uffWordWrap in fmt^.UsedFormattingFields) then           // ... word-wrap
    exit;
  if (uffTextRotation in fmt^.UsedFormattingFields) and        // ... vertical text
     (fmt^.TextRotation <> trHorizontal)
  then
    exit;

  fc := FHeaderCount + FFrozenCols;

  txt := cell^.UTF8Stringvalue;
  if (uffHorAlign in fmt^.UsedFormattingFields) then
    txtalign := fmt^.HorAlignment
  else
    txtalign := haDefault;
  PrepareCanvas(ACol, ARow, AState);
  len := Canvas.TextWidth(txt) + 2*constCellPadding;
  ACol1 := ACol;
  ACol2 := ACol;
  r := GetWorksheetRow(ARow);
  case txtalign of
    haLeft, haDefault:
      // overflow to the right
      while (len > ARect.Right - ARect.Left) and (ACol2 < ColCount-1) do
      begin
        result := true;
        inc(ACol2);
        cell := Worksheet.FindCell(r, GetWorksheetCol(ACol2));
        if (cell <> nil) and (cell^.ContentType <> cctEmpty) then
        begin
          dec(ACol2);
          break;
        end;
        ARect.Right := ARect.Right + ColWidths[ACol2];
      end;
    haRight:
      // overflow to the left
      while (len > ARect.Right - ARect.Left) and (ACol1 > fc) do
      begin
        result := true;
        dec(ACol1);
        cell := Worksheet.FindCell(r, GetWorksheetCol(ACol1));
        if (cell <> nil) and (cell^.ContentType <> cctEmpty) then
        begin
          inc(ACol1);
          break;
        end;
        ARect.Left := ARect.Left - ColWidths[ACol1];
      end;
    haCenter:
      begin
        len := len div 2;
        w0 := (ARect.Right - ARect.Left) div 2;
        w := w0;
        // right part
        while (len > w) and (ACol2 < ColCount-1) do
        begin
          Result := true;
          inc(ACol2);
          cell := Worksheet.FindCell(r, GetWorksheetCol(ACol2));
          if (cell <> nil) and (cell^.ContentType <> cctEmpty) then
          begin
            dec(ACol2);
            break;
          end;
          ARect.Right := ARect.Right + ColWidths[ACol2];
          inc(w, ColWidths[ACol2]);
        end;
        // left part
        w := w0;
        while (len > w) and (ACol1 > fc) do
        begin
          Result := true;
          dec(ACol1);
          cell := Worksheet.FindCell(r, GetWorksheetCol(ACol1));
          if (cell <> nil) and (cell^.Contenttype <> cctEmpty) then
          begin
            inc(ACol1);
            break;
          end;
          ARect.Left := ARect.left - ColWidths[ACol1];
          inc(w, ColWidths[ACol1]);
        end;
      end;
  end;
end;

function TsCustomWorksheetGrid.CellRect(ACol1, ARow1, ACol2, ARow2: Integer): TRect;
var
  cmin, cmax: Integer;
  rmin, rmax: Integer;
  tmp: Integer = 0;
begin
  Result := Rect(0, 0, 0, 0);  // to silence the compiler...
  cmin := Min(ACol1, ACol2);
  cmax := Max(ACol1, ACol2);
  rmin := Min(ARow1, ARow2);
  rmax := Max(ARow1, ARow2);
  if IsRightToLeft then begin
    ColRowToOffset(True, True, cmin, tmp, Result.Right);
    ColRowToOffset(True, True, cmax, Result.Left, tmp);
  end else
  begin
    ColRowToOffset(True, True, cmin, Result.Left, tmp);
    ColRowToOffset(True, True, cmax, tmp, Result.Right);
  end;
  ColRowToOffSet(False, True, rmin, Result.Top, tmp);
  ColRowToOffset(False, True, rmax, tmp, Result.Bottom);
end;

{@@ ----------------------------------------------------------------------------
  Handler for the event OnChangeCell fired by the worksheet when the contents
  or formatting of a cell have changed.
  As a consequence, the grid may have to update the cell.
  Row/Col coordinates are in worksheet units here!

  @param  ASender  Sender of the event OnChangeFont (the worksheet)
  @param  ARow     Row index of the changed cell, in worksheet units!
  @param  ACol     Column index of the changed cell, in worksheet units!
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.ChangedCellHandler(ASender: TObject; ARow, ACol:Cardinal);
begin
  Unused(ASender, ARow, ACol);
  if FLockCount = 0 then Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Handler for the event OnChangeFont fired by the worksheet when the font has
  changed in a cell.
  As a consequence, the grid may have to update the row height.
  Row/Col coordinates are in worksheet units here!

  @param  ASender  Sender of the event OnChangeFont (the worksheet)
  @param  ARow     Row index of the cell with the changed font, in worksheet units!
  @param  ACol     Column index of the cell with the changed font, in worksheet units!
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.ChangedFontHandler(ASender: TObject;
  ARow, ACol: Cardinal);
var
  lRow: PRow;
  gr: Integer; // row index in grid units
begin
  Unused(ASender, ACol);
  if (Worksheet <> nil) then begin
    lRow := Worksheet.FindRow(ARow);
    if lRow = nil then begin
      // There is no row record --> row height changes according to font height
      // Otherwise the row height would be fixed according to the value in the row record.
      gr := GetGridRow(ARow);  // convert row index to grid units
      RowHeights[gr] := CalcAutoRowHeight(gr);
    end;
    Invalidate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Clears the grid contents
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Clear;
begin
  if (Worksheet <> nil) then Worksheet.Clear;
end;

{@@ ----------------------------------------------------------------------------
  Converts a spreadsheet font to a font used for painting (TCanvas.Font).

  @param  sFont  Font as used by fpspreadsheet (input)
  @param  AFont  Font as used by TCanvas for painting (output)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Convert_sFont_to_Font(sFont: TsFont; AFont: TFont);
begin
  fpsVisualUtils.Convert_sFont_to_Font(sFont, AFont);
end;

{@@ ----------------------------------------------------------------------------
  Converts a font used for painting (TCanvas.Font) to a spreadsheet font.

  @param  AFont  Font as used by TCanvas for painting (input)
  @param  sFont  Font as used by fpspreadsheet (output)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Convert_Font_to_sFont(AFont: TFont;
  sFont: TsFont);
begin
  fpsVisualUtils.Convert_Font_to_sFont(AFont, sFont);
end;

{@@ ----------------------------------------------------------------------------
  This is one of the main painting methods inherited from TsCustomGrid. It is
  overridden here to achieve the feature of "frozen" cells which should be
  painted in the same style as normal cells.

  Internally, "frozen" cells are "fixed" cells of the grid. Therefore, it is
  not possible to select any cell within the frozen panes - in contrast to the
  standard spreadsheet applications.

  @param  ACol   Column index of the cell being drawn
  @param  ARow   Row index of the cell beging drawn
  @param  ARect  Rectangle, in grid pixels, covered by the cell
  @param  AState Grid drawing state, as defined by TsCustomGrid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DefaultDrawCell(aCol, aRow: Integer;
  var aRect: TRect; AState: TGridDrawState);
var
  wasFixed: Boolean;
begin
  wasFixed := false;
  if (gdFixed in AState) then
    if ShowHeaders then begin
      if ((ARow < FixedRows) and (ARow > 0) and (ACol > 0)) or
         ((ACol < FixedCols) and (ACol > 0) and (ARow > 0))
      then
        wasFixed := true;
    end else begin
      if (ARow < FixedRows) or (ACol < FixedCols) then
        wasFixed := true;
    end;

  if wasFixed then begin
    AState := AState - [gdFixed];
    Canvas.Brush.Color := clWindow;
    DoPrepareCanvas(ACol, ARow, AState);
  end;

  inherited DefaultDrawCell(ACol, ARow, ARect, AState);

  if wasFixed then begin
    DrawCellGrid(ACol, ARow, ARect, AState);
    AState := AState + [gdFixed];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the column specified.

  @param   AGridCol   Grid index of the column to be deleted
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DeleteCol(AGridCol: Integer);
begin
  if AGridCol < FHeaderCount then
    exit;

  Worksheet.DeleteCol(GetWorksheetCol(AGridCol));
  UpdateColWidths(AGridCol);
end;

{@@ ----------------------------------------------------------------------------
  Deletes the row specified.

  @param  AGridRow   Grid index of the row to be deleted
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DeleteRow(AGridRow: Integer);
begin
  if AGridRow < FHeaderCount then
    exit;

  Worksheet.DeleteRow(GetWorksheetRow(AGridRow));

  // Update following row heights because their index has changed
  UpdateRowHeights(AGridRow);
end;

procedure TsCustomWorksheetGrid.ColRowMoved(IsColumn: Boolean;
  FromIndex,ToIndex: Integer);
begin
  inherited;
  if IsColumn then
    Worksheet.MoveCol(GetWorksheetCol(FromIndex), GetWorksheetCol(ToIndex));
end;

procedure TsCustomWorksheetGrid.CreateHandle;
begin
  inherited;
  Setup;
end;

{@@ ----------------------------------------------------------------------------
  Creates a new empty workbook into which a file will be loaded. Destroys the
  previously used workbook.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.CreateNewWorkbook;
begin
  GetWorkbookSource.CreateNewWorkbook;
  if FReadFormulas then
    WorkbookSource.Options := WorkbookSource.Options + [boReadFormulas] else
    WorkbookSource.Options := Workbooksource.Options - [boReadFormulas];
  SetAutoCalc(FAutoCalc);
end;

{@@ ----------------------------------------------------------------------------
  Is called when a Double-click occurs. Overrides the inherited method to
  react on double click on cell border in row headers to auto-adjust the
  row heights
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DblClick;
var
  oldHeight: Integer;
  gRow: Integer;
begin
  SelectActive := False;
  FGridState := gsNormal;
  if (goRowSizing in Options) and (Cursor = crVSplit) and (FHeaderCount > 0) then
  begin
    if (goDblClickAutoSize in Options) then
    begin
      gRow := GCache.MouseCell.y;
      if CellRect(0, gRow).Bottom - GCache.ClickMouse.y > 0 then dec(gRow);
      oldHeight := RowHeights[gRow];
      AutoAdjustRow(gRow);
      if oldHeight <> RowHeights[gRow] then
        Cursor := crDefault; //ChangeCursor;
    end
  end
  else
    inherited DblClick;
end;

procedure TsCustomWorksheetGrid.DefineProperties(Filer: TFiler);
begin
  //inherited;

  // Don't call inherited, this is where the ColWidths/RowHeights are written
  // to the lfm file - we don't need them, we get them from the workbook!
  Unused(Filer);
end;

procedure TsCustomWorksheetGrid.DoCopyToClipboard;
begin
  WorkbookSource.CopyCellsToClipboard;
end;

procedure TsCustomWorksheetGrid.DoCutToClipboard;
begin
  // This next comment does not seem to be valid any more: Issue handled by eating key in KeyDown
  // Remove for the moment: If TsCopyActions is available this code would be executed twice (and destroy the clipboard)
  WorkbookSource.CutCellsToClipboard;
end;

procedure TsCustomWorksheetGrid.DoPasteFromClipboard;
begin
  // This next comment does not seem to be valid any more: Issue handled by eating key in KeyDown
  // Remove for the moment: If TsPasteActions is available this code would be executed twice
  WorkbookSource.PasteCellsFromClipboard(coCopyCell);
end;

{ Make the cell editor the same size as the edited cell, in particular for
  even for merged cells; otherwise the merge base content would be seen during
  editing at several places. }
procedure TsCustomWorksheetGrid.DoEditorShow;
var
  r1, c1, r2, c2: Cardinal;
  cell: PCell;
  Rct: TRect;
  delta: Integer;
begin
  FOldEditorText := GetCellText(Col, Row);
  inherited;
  if (Worksheet <> nil) and (Editor is TStringCellEditor) then
  begin
    delta := FSelPen.Width div 2;
    cell := Worksheet.FindCell(GetWorksheetRow(Row), GetWorksheetCol(Col));
    if Worksheet.IsMerged(cell) then begin
      Worksheet.FindMergedRange(cell, r1,c1,r2,c2);
      Rct := CellRect(GetGridCol(c1), GetGridRow(r1), GetGridCol(c2), GetGridRow(r2));
    end else
      Rct := CellRect(Col, Row);
    InflateRect(Rct, -delta, -delta);
    inc(Rct.Top);
    if not odd(FSelPen.Width) then dec(Rct.Left);
    Editor.Font.Height := Round(Font.Height * ZoomFactor);
    Editor.SetBounds(Rct.Left, Rct.Top, Rct.Right-Rct.Left-1, Rct.Bottom-Rct.Top-1);
  end;
end;

procedure TsCustomWorksheetGrid.DoOnResize;
begin
  if (csDesigning in ComponentState) and (Worksheet = nil) then
    NewWorkbook(ColCount, RowCount);
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Adjusts the grid's canvas before painting a given cell. Considers
  background color, horizontal alignment, vertical alignment, etc.

  @param  ACol    Column index of the cell being painted
  @param  ARow    Row index of the cell being painted
  @param  AState  Grid drawing state -- see TsCustomGrid.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DoPrepareCanvas(ACol, ARow: Integer;
  AState: TGridDrawState);
var
  ts: TTextStyle;
  lCell: PCell;
  fmt: PsCellFormat;
  r, c: Integer;
  fnt: TsFont;
  isSelected: Boolean;
  fgcolor, bgcolor: TColor;
//  numFmt: TsNumFormatParams;
begin
  GetSelectedState(AState, isSelected);
  Canvas.Font.Assign(Font);
  //Canvas.Font.Height := Round(ZoomFactor * Canvas.Font.Height);
  Canvas.Brush.Bitmap := nil;
  Canvas.Brush.Color := Color;
  ts := Canvas.TextStyle;

  if ShowHeaders then
  begin
    // Formatting of row and column headers
    if ARow = 0 then
    begin
      ts.Alignment := taCenter;
      ts.Layout := tlCenter;
    end else
    if ACol = 0 then
    begin
      ts.Alignment := taRightJustify;
      ts.Layout := tlCenter;
    end;
    if ShowHeaders and ((ACol = 0) or (ARow = 0)) then
      Canvas.Brush.Color := FixedColor
  end;

  if (Worksheet <> nil) and (ARow >= FHeaderCount) and (ACol >= FHeaderCount) then
  begin
    r := ARow - FHeaderCount;
    c := ACol - FHeaderCount;

    fmt := Worksheet.GetPointerToEffectiveCellFormat(r, c);
    lCell := Worksheet.FindCell(r, c);

    //if lCell <> nil then
    //begin

//      fmt := Workbook.GetPointerToCellFormat(lCell^.FormatIndex);
//      numFmt := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);

      // Background color
      if (uffBackground in fmt^.UsedFormattingFields) then
      begin
        if Workbook.FileFormatID = ord(sfExcel2) then
        begin
          CreateFillPattern(FillPatternBitmap, fsGray12, clBlack, Color);
          Canvas.Brush.Style := bsImage;
          Canvas.Brush.Bitmap := FillPatternBitmap;
        end else
        begin
          case fmt^.Background.Style of
            fsNoFill:
              Canvas.Brush.Style := bsClear;
            fsSolidFill:
              begin
                Canvas.Brush.Style := bsSolid;
                Canvas.Brush.Color := fmt^.Background.FgColor and $00FFFFFF;
              end;
            else
              if fmt^.Background.BgColor = scTransparent
                then bgcolor := Color
                else bgcolor := fmt^.Background.BgColor and $00FFFFFF;
              if fmt^.Background.FgColor = scTransparent
                then fgcolor := Color
                else fgcolor := fmt^.Background.FgColor and $00FFFFFF;
              CreateFillPattern(FillPatternBitmap, fmt^.Background.Style, fgColor, bgColor);
              Canvas.Brush.Style := bsImage;
              Canvas.Brush.Bitmap := FillPatternBitmap;
          end;
        end;
      end else
      begin
        Canvas.Brush.Style := bsSolid;
        Canvas.Brush.Color := Color;
      end;

      // Font
      if (lcell <> nil) and Worksheet.HasHyperlink(lCell) then
        fnt := Workbook.GetHyperlinkFont
      else
        fnt := Workbook.GetDefaultFont;
      if (uffFont in fmt^.UsedFormattingFields) then
        fnt := Workbook.GetFont(fmt^.FontIndex);

      Convert_sFont_to_Font(fnt, Canvas.Font);
      Canvas.Font.Height := Round(ZoomFactor * Canvas.Font.Height);

      // Wordwrap, text alignment and text rotation are handled by "DrawTextInCell".
    //end;
  end;

  if IsSelected then
    Canvas.Brush.Color := CalcSelectionColor(Canvas.Brush.Color, 16);

  Canvas.TextStyle := ts;

  inherited DoPrepareCanvas(ACol, ARow, AState);
end;

procedure TsCustomWorksheetGrid.DragDrop(Source: TObject; X, Y: Integer);
var
  i, j: Integer;
  sel: TsCellRange;
  srccell, destcell: PCell;
  r: LongInt = 0;
  c: LongInt = 0;
  dr, dc: LongInt;
  dragMove: Boolean;
begin
  Unused(X, Y);

  inherited;

  if not ((goEditing in Options)) or (not FAllowDragAndDrop) then
    Exit;

  // Offset of col/row coordinates from source to destination cell
  MouseToCell(X,Y, c, r);
  dr := r - FDragStartRow;
  dc := c - FDragStartCol;

  // Copy cells or only move them?
  dragMove := not (ssCtrl in GetKeyShiftState);

  // Copy cells to destination and delete the source cells if required.
  for sel in Worksheet.GetSelection do
  begin
    for r := sel.Row1 to sel.Row2 do
      for c := sel.Col1 to sel.Col2 do
      begin
        srccell := Worksheet.FindCell(r, c);
        if Worksheet.IsMerged(srccell) then
          srccell := Worksheet.FindMergeBase(srccell);
        if srccell <> nil then begin
          destcell := Worksheet.GetCell(r + dr, c + dc);
          Worksheet.CopyCell(srccell, destcell);
          if dragMove then
            Worksheet.DeleteCell(srccell);
        end;
      end;
  end;
end;

procedure TsCustomWorksheetGrid.DragOver(ASource: TObject; X, Y: Integer;
  AState: TDragState; var Accept: Boolean);
var
  destcell: PCell;
  sc, sr: Integer;
  gc, gr: Integer;
  dc, dr: Integer;
  sel: TsCellRange;
  dragMove: Boolean;
begin
  inherited;
  Unused(AState);

  if FAllowDragAndDrop and (ASource = self) and (goEditing in Options) then
  begin
    MouseToCell(X,Y, gc, gr);

    // Don't drop over over the header cells
    if (gc < FHeaderCount) or (gr < FHeaderCount) then
      exit;

    // Find dragged selection rectangle
    dc := gc - FDragStartCol;
    dr := gr - FDragStartRow;
    FDragSelection := Selection;
    OffsetRect(FDragSelection, dc, dr);

    // Draw dragged selection rectangle
    if (FOldDragStartRow <> gr) or (FOldDragStartCol <> gc) then
      Invalidate;
    FOldDragStartRow := gr;
    FOldDragStartCol := gc;

    // Change mouse cursor: Copy or Move
    dragMove := not (ssCtrl in GetKeyShiftState);
    if dragMove then
      DragCursor := crDrag
    else
      DragCursor := crDragCopy;

    if Worksheet.IsProtected then
      // Allow drop only if no destination cell is locked
      for sel in Worksheet.GetSelection do
        for sr := sel.Row1 to sel.Row2 do
          for sc := sel.Col1 to sel.Col2 do
          begin
            destcell := Worksheet.FindCell(sr + dr, sc + dc);
            if (cpLockCell in Worksheet.ReadCellProtection(destcell)) then
              exit;
          end;
    Accept := true;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This method is inherited from TsCustomGrid, but is overridden here in order
  to paint the cell borders and the selection rectangle.
  Both features can extend into the neighboring cells and thus would be clipped
  at the cell borders by the standard painting mechanism. At the time when
  DrawAllRows is called, however, clipping at cell borders is no longer active.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawAllRows;
var
  cliprect: TRect;
  rgn: HRGN;
  //tmp: Integer = 0;
  //fc, fr: Integer;
  TL: TPoint;
begin
  inherited;

  FTopLeft := CalcTopLeft(false);

  if (FrozenRows > 0) or (FrozenCols > 0) then
    DrawFrozenPanes;

  // Set cliprect for scrollable grid area
  cliprect := ClientRect;
  TL := CalcTopLeft(false);
  if IsRightToLeft then
    cliprect.Right := TL.X + 1
  else
    cliprect.Left := TL.X - 1;
  cliprect.Top := TL.Y;

  // Paint cell borders, selection rectangle, images and frozen-pane-borders
  // into this clipped area
  rgn := CreateRectRgn(cliprect.Left, cliprect.top, cliprect.Right, cliprect.Bottom);
  try
    SelectClipRgn(Canvas.Handle, Rgn);
    DrawCellBorders;
    DrawSelection;
    DrawDragSelection;
    DrawImages(DRAW_NON_FROZEN);
  //  DrawFrozenPaneBorders(clipRect);
  finally
    DeleteObject(rgn);
  end;
end;

procedure TsCustomWorksheetGrid.DrawFrozenPanes;
var
  cliprect, R: TRect;
  rgn: HRGN;
  tmp: Integer = 0;
  fc, fr: Integer;
begin
  if Worksheet = nil then
    exit;

  Canvas.SaveHandleState;
  try
    // Avoid painting into header cells.
    R := ClientRect;
    if HeaderCount > 0 then begin
      ColRowToOffset(false, True, 0, tmp, R.Top);
      if IsRightToLeft then
        ColRowToOffset(true, True, 0, R.Right, tmp)
      else
        ColRowToOffset(true, True, 0, tmp, R.Left);
    end;

    fr := FHeaderCount + FFrozenRows;
    fc := FHeaderCount + FFrozenCols;

    // Paint cell border in frozen rows
    if fr > 0 then begin
      if IsRightToLeft then
        clipRect := Rect(ClientRect.Left, ClientRect.Top, FTopLeft.X, FTopLeft.Y)
      else
        cliprect := Rect(FTopLeft.X, ClientRect.Top, ClientRect.Right, FTopLeft.Y);
      rgn := CreateRectRgn(cliprect.Left, cliprect.Top, cliprect.Right, cliprect.Bottom);
      try
        SelectClipRgn(Canvas.Handle, rgn);
        DrawCellBorders(DRAW_FROZEN_ROWS);
        DrawImages(DRAW_FROZEN_ROWS);
        DrawFrozenPaneBorder(cliprect.Left, clipRect.Right, cliprect.Bottom-1, true);
      finally
        DeleteObject(rgn);
      end;
    end;

    // Paint cell border in frozen columns
    if fc > 0 then begin
      if IsRightToLeft then
        cliprect := Rect(FTopLeft.X, FTopLeft.Y, ClientRect.Right, ClientRect.Bottom)
      else
        cliprect := Rect(ClientRect.Left, FTopLeft.Y, FTopLeft.X, ClientRect.Bottom);
      rgn := CreateRectRgn(cliprect.Left, cliprect.Top, cliprect.Right, cliprect.Bottom);
      try
        SelectClipRgn(Canvas.Handle, rgn);
        DrawCellBorders(DRAW_FROZEN_COLS);
        DrawImages(DRAW_FROZEN_COLS);
        if IsRightToLeft then
          DrawFrozenPaneBorder(cliprect.Top, cliprect.Bottom, cliprect.Left+1, false)
        else
          DrawFrozenPaneBorder(cliprect.Top, cliprect.Bottom, cliprect.Right-1, false);
      finally
        DeleteObject(rgn);
      end;
    end;

    // Paint intersection of frozen cols and frozen rows
    if (fr > 0) and (fc > 0) then begin
      if IsRightToLeft then
        cliprect := Rect(FTopLeft.X, ClientRect.Top, ClientRect.Right, FTopLeft.Y)
      else
        cliprect := Rect(ClientRect.Left, ClientRect.Top, FTopLeft.X, FTopLeft.Y);
      rgn := CreateRectRgn(clipRect.Left, cliprect.Top, cliprect.Right, cliprect.Bottom);
      try
        SelectClipRgn(Canvas.Handle, rgn);
        DrawCellBorders(DRAW_FROZEN_CORNER);
        DrawImages(DRAW_FROZEN_CORNER);
      finally
        DeleteObject(rgn);
      end;
    end;
  finally
    Canvas.RestoreHandleState;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Draws the borders of all cells. Calls DrawCellBorders for each individual cell.
  AGridPart denotes where the cells are painted:
  0 = normal grid area
  1 = FrozenRows
  2 = FrozenCols
  3 = Top-left corner where FrozenCols and FrozenRows intersect
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawCellBorders(AGridPart: Integer = 0);
var
  cell, base: PCell;
  gc, gr: Integer;
  sr1, sc1, sr2, sc2: Cardinal;
  rect: TRect;
  cellHasBorder: Boolean;
begin
  if Worksheet = nil then
    exit;

  case AGridPart of
    0: begin
         sr1 := GetWorksheetRow(GCache.VisibleGrid.Top);
         sc1 := GetWorksheetCol(GCache.VisibleGrid.Left);
         sr2 := GetWorksheetRow(GCache.VisibleGrid.Bottom);
         sc2 := GetWorksheetCol(GCache.VisibleGrid.Right);
       end;
    1: begin
         sr1 := 0;
         sr2 := FFrozenRows - 1;
         sc1 := FFrozenCols - 1;
         sc2 := GetWorksheetCol(GCache.VisibleGrid.Right);
       end;
    2: begin
         sc1 := 0;
         sc2 := FFrozenCols - 1;
         sr1 := FFrozenRows - 1;
         sr2 := GetWorksheetRow(GCache.VisibleGrid.Bottom);
       end;
    3: begin
         sc1 := 0;
         sc2 := FFrozenCols - 1;
         sr1 := 0;
         sr2 := FFrozenRows - 1;
       end;
  end;
  if sr1 = UNASSIGNED_ROW_COL_INDEX then sr1 := 0;
  if sc1 = UNASSIGNED_ROW_COL_INDEX then sc1 := 0;

  for cell in Worksheet.Cells.GetRangeEnumerator(sr1, sc1, sr2, sc2) do
  begin
    if Worksheet.IsMerged(cell) then
    begin
      base := Worksheet.FindMergeBase(cell);
      cellHasBorder := uffBorder in Worksheet.ReadUsedFormatting(base);
    end else
      cellHasBorder := uffBorder in Worksheet.ReadUsedFormatting(cell);
    if cellHasBorder then
    begin
      gc := GetGridCol(cell^.Col);
      gr := GetGridRow(cell^.Row);
      rect := CellRect(gc, gr);
      DrawCellBorders(gc, gr, rect, cell);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Draws the border lines around a given cell. Note that when this procedure is
  called the output is clipped by the cell rectangle, but thick and double
  border styles extend into the neighboring cell. Therefore, these border lines
  are drawn in parts.

  @param  ACol   Column Index
  @param  ARow   Row index
  @param  ARect  Rectangle in pixels occupied by the cell.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawCellBorders(ACol, ARow: Integer;
  ARect: TRect; ACell: PCell);
const
  drawHor = 0;
  drawVert = 1;
  drawDiagUp = 2;
  drawDiagDown = 3;

  procedure DrawBorderLine(ACoord: Integer; ARect: TRect; ADrawDirection: Byte;
    ABorderStyle: TsCellBorderStyle);
  const
    //TsLineStyle = (
    //  lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair,
    //  lsMediumDash, lsDashDot, lsMediumDashDot, lsDashDotDot, lsMediumDashDotDot,
    //  lsSlantDashDot);
    PEN_STYLES: array[TsLineStyle] of TPenStyle =
      (psSolid, psSolid, psDash, psDot, psSolid, psSolid, psDot,
       psDash, psDashDot, psDashDot, psDashDotDot, psDashDotDot,
       psDashDot);
    PEN_WIDTHS: array[TsLineStyle] of Integer =
      (1, 2, 1, 1, 3, 1, 1,
       2, 1, 2, 1, 2,
       2);
  var
    deltax, deltay: Integer;
    angle: Double;
    savedCosmetic: Boolean;
  begin
    savedCosmetic := Canvas.Pen.Cosmetic;
    Canvas.Pen.Style := PEN_STYLES[ABorderStyle.LineStyle];
    Canvas.Pen.Width := PEN_WIDTHS[ABorderStyle.LineStyle];
    Canvas.Pen.Color := ABorderStyle.Color and $00FFFFFF;
    Canvas.Pen.EndCap := pecSquare;
    if ABorderStyle.LineStyle = lsHair then
      Canvas.Pen.Cosmetic := false;
    // Painting
    case ABorderStyle.LineStyle of
      lsThin, lsMedium, lsThick, lsDotted, lsDashed, lsDashDot, lsDashDotDot,
      lsMediumDash, lsMediumDashDot, lsMediumDashDotDot, lsSlantDashDot, lsHair:
        case ADrawDirection of
          drawHor     : Canvas.Line(ARect.Left, ACoord, ARect.Right, ACoord);
          drawVert    : Canvas.Line(ACoord, ARect.Top, ACoord, ARect.Bottom);
          drawDiagUp  : Canvas.Line(ARect.Left, ARect.Bottom, ARect.Right, ARect.Top);
          drawDiagDown: Canvas.Line(ARect.Left, ARect.Top, ARect.Right, ARect.Bottom);
        end;
      {
      lsHair:
        case ADrawDirection of
          drawHor     : DrawHairLineHor(Canvas, ARect.Left, ARect.Right, ACoord);
          drawVert    : DrawHairLineVert(Canvas, ACoord, ARect.Top, ARect.Bottom);
          drawDiagUp  : ;
          drawDiagDown: ;
        end;
      }
      lsDouble:
        case ADrawDirection of
          drawHor:
            begin
              Canvas.Line(ARect.Left, ACoord-1, ARect.Right, ACoord-1);
              Canvas.Line(ARect.Left, ACoord+1, ARect.Right, ACoord+1);
              Canvas.Pen.Color := Color;
              Canvas.Line(ARect.Left, ACoord, ARect.Right, ACoord);
            end;
          drawVert:
            begin
              Canvas.Line(ACoord-1, ARect.Top, ACoord-1, ARect.Bottom);
              Canvas.Line(ACoord+1, ARect.Top, ACoord+1, ARect.Bottom);
              Canvas.Pen.Color := Color;
              Canvas.Line(ACoord, ARect.Top, ACoord, ARect.Bottom);
            end;
          drawDiagUp:
            begin
              if ARect.Right = ARect.Left then
                angle := pi/2
              else
                angle := arctan((ARect.Bottom-ARect.Top) / (ARect.Right-ARect.Left));
              deltax := Max(1, round(1.0 / sin(angle)));
              deltay := Max(1, round(1.0 / cos(angle)));
              Canvas.Line(ARect.Left, ARect.Bottom-deltay-1, ARect.Right-deltax, ARect.Top-1);
              Canvas.Line(ARect.Left+deltax, ARect.Bottom-1, ARect.Right, ARect.Top+deltay-1);
            end;
          drawDiagDown:
            begin
              if ARect.Right = ARect.Left then
                angle := pi/2
              else
                angle := arctan((ARect.Bottom-ARect.Top) / (ARect.Right-ARect.Left));
              deltax := Max(1, round(1.0 / sin(angle)));
              deltay := Max(1, round(1.0 / cos(angle)));
              Canvas.Line(ARect.Left, ARect.Top+deltay-1, ARect.Right-deltax, ARect.Bottom-1);
              Canvas.Line(ARect.Left+deltax, ARect.Top-1, ARect.Right, ARect.Bottom-deltay-1);
            end;
        end;
    end;
    Canvas.Pen.Cosmetic := savedCosmetic;
  end;

var
  bs: TsCellBorderStyle;
  fmt: PsCellFormat;
  r1, c1, r2, c2: Cardinal;
begin
  if Assigned(Worksheet) then begin
    if Worksheet.IsMergeBase(ACell) then begin
      Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
      ARect := CellRect(GetGridCol(c1), GetGridRow(r1), GetGridCol(c2), GetGridRow(r2));
    end;

    // Left border
    if GetBorderStyle(ACol, ARow, -1, 0, ACell, bs) then
      if IsRightToLeft then
        DrawBorderLine(ARect.Right, ARect, drawVert, bs)
      else
        DrawBorderLine(ARect.Left-ord(not IsRightToLeft), ARect, drawVert, bs);
    // Right border
    if GetBorderStyle(ACol, ARow, +1, 0, ACell, bs) then
      if IsRightToLeft then
        DrawBorderLine(ARect.Left, ARect, drawVert, bs)
      else
        DrawBorderLine(ARect.Right-ord(not IsRightToLeft), ARect, drawVert, bs);
    // Top border
    if GetBorderstyle(ACol, ARow, 0, -1, ACell, bs) then
      DrawBorderLine(ARect.Top-1, ARect, drawHor, bs);
    // Bottom border
    if GetBorderStyle(ACol, ARow, 0, +1, ACell, bs) then
      DrawBorderLine(ARect.Bottom-1, ARect, drawHor, bs);

    if ACell <> nil then begin
      fmt := Worksheet.GetPointerToEffectiveCellFormat(ACell);
      // Diagonal up
      if cbDiagUp in fmt^.Border then begin
        bs := fmt^.Borderstyles[cbDiagUp];
        if IsRightToLeft then
          DrawBorderLine(0, ARect, drawDiagDown, bs)
        else
          DrawBorderLine(0, ARect, drawDiagUp, bs);
      end;
      // Diagonal down
      if cbDiagDown in fmt^.Border then begin
        bs := fmt^.BorderStyles[cbDiagDown];
        if IsRightToLeft then
          DrawBorderLine(0, ARect, drawDiagUp, bs)
        else
          DrawborderLine(0, ARect, drawDiagDown, bs);
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Method inherited method from TCustomGrid. Is overridden here to avoid painting
  of the border of frozen cells in black under some circumstances.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawCellGrid(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
begin
  if (TitleStyle <> tsNative) and (gdFixed in AState) and
     {DisplayFixedColRow and} ((FFrozenCols > 0) or (FFrozenRows > 0)) then
  begin
    // Draw default cell borders only in the header cols/rows.
    // If there are frozen cells they would get a black border, so we don't
    // draw their borders here - they are drawn by "DrawRow" anyway.
    if ((ACol=0) or (ARow = 0)) and DisplayFixedColRow then
      inherited;
  end else
    inherited;
end;

{@@ ----------------------------------------------------------------------------
  Draws the red rectangle in the upper right corner of a cell to indicate that
  this cell contains a popup comment
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawCommentMarker(ARect: TRect);
const
  COMMENT_SIZE = 7;
var
  P: Array[0..3] of TPoint;
begin
  Canvas.Brush.Color := clRed;
  Canvas.Brush.Style := bsSolid;
  Canvas.Pen.Style := psClear;
  if IsRightToLeft then
  begin
    P[0] := Point(ARect.Left, ARect.Top);
    P[1] := Point(ARect.Left + COMMENT_SIZE, ARect.Top);
    P[2] := Point(ARect.Left, ARect.Top + COMMENT_SIZE);
  end else
  begin
    P[0] := Point(ARect.Right, ARect.Top);
    P[1] := Point(ARect.Right - COMMENT_SIZE, ARect.Top);
    P[2] := Point(ARect.Right, ARect.Top + COMMENT_SIZE);
  end;
  P[3] := P[0];
  Canvas.Polygon(P);
end;

{@@ ----------------------------------------------------------------------------
  This procedure is responsible for painting the focus rectangle. We don't want
  the red dashed rectangle here, but prefer the thick Excel-like black border
  line. This new focus rectangle is drawn by the method DrawSelection because
  the thick Excel border reaches into adjacent cells.

  @param   ACol   Grid column index of the focused cell
  @param   ARow   Grid row index of the focused cell
  @param   ARect  Rectangle in pixels covered by the focused cell
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawFocusRect(aCol, aRow: Integer; ARect: TRect);
begin
  Unused(ACol, ARow, ARect);
  // Nothing do to
end;

{@@ ----------------------------------------------------------------------------
  Draws a solid line along the borders of frozen panes.

  @param  AStart  Start coordinate of the pane border line
  @param  AEnd    End coordinate of the pane border line
  @param  ACoord  other coordinate of the border line
                  (y if horizontal, x if vertical)
  @param  IsHor   Determines whether a horizontal or vertical line is drawn and,
                  thus, how AStart, AEnd and ACoord are interpreted.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawFrozenPaneBorder(AStart, AEnd, ACoord: Integer;
  IsHor: Boolean);
var
  delta: Integer;
begin
  if (IsHor and (FFrozenRows = 0)) or (not IsHor and (FFrozenCols = 0)) then
    exit;

  if FFrozenBorderPen.Style = psClear then
    exit;

  delta := (FFrozenBorderPen.Width - 1) div 2;

  Canvas.Pen.Assign(FFrozenBorderPen);
  if IsHor then
    Canvas.Line(AStart, ACoord-delta, AEnd, ACoord-delta)
  else
  if IsRightToLeft then
    Canvas.Line(ACoord+delta, AStart, ACoord+delta, AEnd)
  else
    Canvas.Line(ACoord-delta, AStart, ACoord-delta, AEnd);
end;

{@@ ----------------------------------------------------------------------------
  Draws the embedded images of the worksheet. Is called at the end of the
  painting process.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawImages(AGridPart: Integer);

  procedure CalcClipRect(var ARect: TRect);
  var
    headerTL: TPoint;
  begin
    ARect := ClientRect;
    headerTL := CalcTopLeft(true);
    case AGridPart of
      DRAW_NON_FROZEN:
        begin
          if IsRightToLeft then
            ARect.Right := FTopLeft.X
          else
            ARect.Left := FTopLeft.X;
          ARect.Top := FTopLeft.Y;
        end;
      DRAW_FROZEN_ROWS:
        begin
          if IsRightToLeft then
            ARect.Right := FTopLeft.X
          else
            ARect.Left := FTopLeft.X;
          ARect.Top := headerTL.Y;
          ARect.Bottom := FTopLeft.Y;
        end;
      DRAW_FROZEN_COLS:
        begin
          if IsRightToLeft then
          begin
            ARect.Left := FTopLeft.X;
            ARect.Right := headerTL.X;
          end else
          begin
            ARect.Left := headerTL.X;
            ARect.Right := FTopLeft.X;
          end;
          ARect.Top := FTopLeft.Y;
        end;
      DRAW_FROZEN_CORNER:
        begin
          if IsRightToLeft then
          begin
            ARect.Left := FTopLeft.X;
            ARect.Right := headerTL.X;
          end else
          begin
            ARect.Left := headerTL.X;
            ARect.Right := FTopLeft.X;
          end;
          ARect.Top := headerTL.Y;
          ARect.Bottom := FTopLeft.Y;
        end;
    end;
  end;

  // Offset to convert relative to absolute row/col coordinates for ColRowToOffset
  procedure GetScrollOffset(out ARowDelta, AColDelta: Integer);
  var
    tmp: Integer;
    x: Integer = 0;
    y: Integer = 0;
  begin
    AColDelta := 0;
    if FrozenCols > 0 then begin
      if IsRightToLeft then begin
        tmp := LeftCol;
        ColRowToOffset(true, false, LeftCol, AColDelta, tmp); //tmp, AColDelta);
        ColRowToOffset(true, true, LeftCol, tmp, x);
        dec(AColDelta, ClientWidth - x);
      end else
      begin
        ColRowToOffset(true, false, LeftCol, AColDelta, tmp);
        ColRowToOffset(true, true, LeftCol, x, tmp);
        dec(AColDelta, x);
      end;
    end;

    ARowDelta := 0;
    if FrozenRows > 0 then begin
      ColRowToOffset(false, false, TopRow, ARowDelta, tmp);
      ColRowToOffset(false, true, TopRow, y, tmp);
      dec(ARowDelta, y);
    end;
  end;

  function GetImageRect(img: PsImage; AWidth, AHeight: Integer;
    ARowDelta, AColDelta: Integer): TRect;
  var
    tmp: Integer = 0;
    gcol, grow: Integer;
    relativeX, relativeY: Boolean;
  begin
    Result := Rect(0, 0, 0, 0);  // To silence the compiler

    grow := GetGridRow(img^.row);
    gcol := GetGridCol(img^.Col);

    case AGridPart of
      DRAW_NON_FROZEN:
        begin
          relativeX := (FrozenCols = 0);
          relativeY := (FrozenRows = 0);
        end;
      DRAW_FROZEN_COLS:
        begin
          relativeX := true;
          relativeY := not ((integer(img^.Row) < FrozenRows) and
                            (integer(img^.Col) < FrozenCols));
        end;
      DRAW_FROZEN_ROWS:
        begin
          relativeX := not ((integer(img^.Row) < FrozenRows) and
                           (integer(img^.Col) < FrozenCols));
          relativeY := true;
        end;
      DRAW_FROZEN_CORNER:
        begin
          relativeX := true;
          relativeY := true;
        end;
    end;

    if IsRightToLeft then begin
      if not relativeX then
        ColRowToOffset(true, false, gcol, Result.Right, tmp)
      else
        ColRowToOffset(true, true, gcol, tmp, Result.Right);
      if not relativeX then
        Result.Right := ClientWidth - Result.Right + AColDelta;
      Result.Left := Result.Right - AWidth;
    end else
    begin
      ColRowToOffset(true, relativeX, gcol, Result.Left, tmp);
      if not relativeX then
        dec(Result.Left, AColDelta);
      Result.Right := Result.Left + AWidth;
    end;

    ColRowToOffset(false, relativeY, grow, Result.Top, tmp);
    if not relativeY then
      dec(Result.Top, ARowDelta);
    Result.Bottom := Result.Top + AHeight;

    if IsRightToLeft then
      OffsetRect(Result, -ToPixels(img^.OffsetX), ToPixels(img^.OffsetY))
    else
      OffsetRect(Result, ToPixels(img^.OffsetX), ToPixels(img^.OffsetY));
  end;

var
  i: Integer;
  img: PsImage;
  obj: TsEmbeddedObj;
  clipArea: TRect = (Left:0; Top:0; Right:0; Bottom:0);
  imgRect: TRect;
  w, h: Integer;
  coloffs, rowoffs: Integer;
  pic: TPicture;
  rgn: HRGN;
  fc, fr: Integer;
  R: TRect = (Left:0; Top:0; Right: 0; Bottom: 0);
begin
  if Worksheet.GetImageCount = 0 then
    exit;

  CalcClipRect(clipArea);
  GetScrollOffset(rowOffs, colOffs);
  fc := FHeaderCount + FFrozenCols;
  fr := FHeaderCount + FFrozenRows;
                           (*
  Canvas.SaveHandleState;
  try
    // Draw bitmap over grid. Take care of clipping.
    InterSectClipRect(Canvas.Handle,
      clipArea.Left, clipArea.Top, clipArea.Right, clipArea.Bottom);
                             *)
    for i := 0 to Worksheet.GetImageCount-1 do begin
      img := Worksheet.GetPointerToImage(i);
      obj := Workbook.GetEmbeddedObj(img^.Index);

      // Frozen part of the grid draw only images which are anchored there.
      case AGridPart of
        DRAW_NON_FROZEN:
          ;
        DRAW_FROZEN_ROWS:
          if (integer(img^.Row) >= fr) then Continue;
        DRAW_FROZEN_COLS:
          if (integer(img^.Col) >= fc) then Continue;
        DRAW_FROZEN_CORNER:
          if (integer(img^.Row) >= fr) or (integer(img^.Col) >= fc) then Continue;
      end;

      // Size of image and its position
      w := ToPixels(obj.ImageWidth * img^.ScaleX);
      h := ToPixels(obj.ImageHeight * img^.ScaleY);
      imgRect := GetImageRect(img, w, h, rowoffs, coloffs);

      // Nothing to do if image is outside the visible grid area
      if not IntersectRect(R, clipArea, imgRect) then
        continue;

      // If not yet done load image stream into bitmap and scale to required size
      if img^.Bitmap = nil then begin
        img^.Bitmap := TBitmap.Create;
        TBitmap(img^.Bitmap).SetSize(w, h);
        TBitmap(img^.Bitmap).PixelFormat := pf32Bit;
        TBitmap(img^.Bitmap).Transparent := true;
      end;

      pic := TPicture.Create;
      try
        obj.Stream.Position := 0;
        pic.LoadFromStream(obj.Stream);
        if pic.Bitmap <> nil then
          TBitmap(img^.Bitmap).Canvas.StretchDraw(Rect(0, 0, w, h), pic.Bitmap)
        else if pic.Graphic <> nil then
          TBitmap(img^.Bitmap).Canvas.StretchDraw(Rect(0, 0, w, h), pic.Graphic);
      finally
        pic.Free;
      end;

      // Draw the bitmap
      rgn := CreateRectRgn(clipArea.Left, clipArea.Top, clipArea.Right, clipArea.Bottom);
      try
        SelectClipRgn(Canvas.Handle, rgn);
        Canvas.Draw(imgRect.Left, imgRect.Top, TBitmap(img^.Bitmap));
      finally
        DeleteObject(rgn);
      end;
    end;
           {
  finally
    Canvas.RestoreHandleState;
  end;      }

end;

{@@ ----------------------------------------------------------------------------
  Draws a complete row of cells. Is mostly duplicated from Grids.pas, but adds
  code for merged cells and overflow text, the section for drawing the default
  focus rectangle is removed.

  @param  ARow  Index of the row to be drawn (in grid coordinates)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawRow(ARow: Integer);
var
  gr, gc, gcLast: Integer;    // grid row/column
  fc: Integer;  // Fixed columns (= header column + frozen columns)
  rct, row_rct, header_rct: TRect;
  clipArea: TRect;
begin
  clipArea := Canvas.ClipRect;

  // Upper and lower bounds for this row
  row_rct := Rect(clipArea.Left, 0, clipArea.Right, 0);
  ColRowToOffSet(False, True, ARow, row_rct.Top, row_rct.Bottom);

  // Rectangle covering the fixed row headers (but not the frozen cells)
  header_rct := Rect(0, 0, 0, 0);
  header_rct.Top := row_rct.Top;
  header_rct.Bottom := row_rct.Bottom;
  if HeaderCount > 0 then
    ColRowToOffset(true, true, 0, header_rct.Left, header_rct.Right);

  // Don't draw rows outside the clipping area
  if (row_rct.Top >= row_rct.Bottom) or not VerticalIntersect(row_rct, clipArea) then
  begin
    {$IFDEF DbgVisualChange}
    DebugLn('Drawrow: Skipped row: ', IntToStr(aRow));
    {$ENDIF}
    exit;
  end;

  // Count of non-scrolling columns
  fc := FHeaderCount + FFrozenCols;

  // (1) Draw data columns in this row (non-fixed part)
  with GCache.VisibleGrid do
  begin
    gcLast := Right;
    gc := Left;
    rct := row_rct;
    if IsRightToLeft then
      rct.Right := FTopLeft.X
    else
      rct.Left := FTopLeft.X;
    InternalDrawRow(ARow, gc, gcLast, rct);
  end;

  // (2) Draw fixed columns consisting of header columns and frozen cells
  gr := ARow;
  // (2a) Draw header column
  if FHeaderCount > 0 then begin
    FDrawingCell := nil;
    gc := 0;
    InternalDrawCell(gc, gr, header_rct, header_rct, [gdFixed]);
  end;

  // (2b) Draw frozen cells
  if FFrozenCols > 0 then begin
    rct := row_rct;
    if IsRightToLeft then
      rct.Left := FTopLeft.X
    else
      rct.Right := FTopLeft.X;
    InternalDrawRow(ARow, FHeaderCount, fc-1, rct);
  end;
end;

procedure TsCustomWorksheetGrid.DrawDragSelection;
begin
  if Assigned(DragManager) and DragManager.IsDragging then
    InternalDrawSelection(FDragSelection, false);
end;

{@@ ----------------------------------------------------------------------------
  Draws the selection rectangle around selected cells.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawSelection;
begin
  InternalDrawSelection(Selection, true);
end;

procedure TsCustomWorksheetGrid.InternalDrawSelection(ASel: TGridRect;
  IsNormalSelection: boolean);
var
  R: TRect;
  cell: PCell;
  r1,c1,r2,c2: Cardinal;
  delta: Integer;
  savedPenMode: TPenMode;
  penwidth: Integer;
  P: array[0..9] of TPoint;
begin
  if Worksheet = nil then
    exit;

  // Selected cell
  cell := Worksheet.FindCell(GetWorksheetRow(ASel.Top), GetWorksheetCol(ASel.Left));
  if Worksheet.IsMerged(cell) then
  begin
    Worksheet.FindMergedRange(cell, r1,c1,r2,c2);
    R := CellRect(GetGridCol(c1), GetGridRow(r1), GetGridCol(c2), GetGridRow(r2));
  end else
    R := CellRect(ASel.Left, ASel.Top, ASel.Right, ASel.Bottom);

  if IsNormalSelection then begin
    // Fine-tune position of selection rect
    if odd(FSelPen.Width) then delta := -1 else delta := 0;
    inc(R.Top, delta);
    if IsRightToLeft then begin
      if not odd(FSelPen.Width) then
        OffsetRect(R, 1, 0) else
        inc(R.Right, 1);
    end else
      inc(R.Left, delta);
    if (ASel.Top = TopRow) then
      inc(R.Top);
    if ASel.Left = LeftCol then begin
      if IsRightToLeft then
        dec(R.Right)
      else
        inc(R.Left);
    end;

    // Set up the canvas
    savedPenMode := Canvas.Pen.Mode;
    Canvas.Pen.Assign(FSelPen);
    if UseXORFeatures then begin
      Canvas.Pen.Color := clWhite;
      Canvas.Pen.Mode := pmXOR;
    end;
    Canvas.Brush.Style := bsClear;
    // Paint
    Canvas.Rectangle(R);
    // Restore canvas.
    Canvas.Pen.Mode := savedPenMode;
  end
  else
  begin
    // Selection during dragging: draw a dotted filled outline
    delta := 2;
    // outer rectangle
    P[0] := Point(R.Left - delta, R.Top - delta);
    P[1] := Point(R.Right + delta, R.Top - delta);
    P[2] := Point(R.Right + delta, R.Bottom + delta);
    P[3] := Point(R.Left - delta, R.Bottom + delta);
    P[4] := P[0];
    // inner rectangle
    P[5] := Point(R.Left + delta, R.Top + delta);
    P[6] := Point(R.Left + delta, R.Bottom - delta);
    P[7] := Point(R.Right - delta, R.Bottom - delta);
    P[8] := Point(R.Right - delta, R.Top + delta);
    P[9] := P[5];
    Canvas.Pen.style := psClear;
    Canvas.Brush.Style := bsImage;
    Canvas.Brush.Bitmap := DragBorderBitmap;
//    Canvas.Brush.Color := clblack;
    Canvas.Polygon(P);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Draws the cell text. Calls "GetCellText" to determine the text for the cell.
  Takes care of horizontal and vertical text alignment, text rotation and
  text wrapping.

  @param  ACol   Grid column index of the cell
  @param  ARow   Grid row index of the cell
  @param  ARect  Rectangle in pixels occupied by the cell
  @param  AState Drawing state of the grid -- see TCustomGrid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.DrawTextInCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
var
  ts: TTextStyle;
  txt: String;
  wrapped: Boolean;
  horAlign: TsHorAlignment;
  vertAlign: TsVertAlignment;
  txtRot: TsTextRotation;
  fntIndex: Integer;
  lCell: PCell;
  fmt: PsCellFormat;
  numfmt: TsNumFormatParams;
  numFmtColor: TColor;
  sidx: Integer;   // number format section index
  RTL: Boolean;
begin
  if (Worksheet = nil) then
    exit;

  if (ACol < FHeaderCount) or (ARow < FHeaderCount) then
    lCell := nil
  else begin
    if FDrawingCell = nil then
      lCell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol))
    else
      lCell := FDrawingCell
  end;

  // Header
  if (lCell = nil) and ShowHeaders and ((ACol = 0) or (ARow = 0)) then
  begin
    ts := Canvas.TextStyle;
    ts.Alignment := taCenter;
    ts.Layout := tlCenter;
    ts.Opaque := false;
    Canvas.TextStyle := ts;
    inherited DrawCellText(aCol, aRow, aRect, aState, GetCellText(ACol,ARow));
    exit;
  end;

  // Cells
  if lCell = nil then
    exit;

  txt := GetCellText(GetGridRow(lCell^.Col), GetGridCol(lCell^.Row));
  if txt = '' then
    exit;

//  fmt := Workbook.GetPointerToCellFormat(lCell^.FormatIndex);
  fmt := Worksheet.GetPointerToEffectiveCellFormat(lCell);
  wrapped := (uffWordWrap in fmt^.UsedFormattingFields) or (fmt^.TextRotation = rtStacked);
  RTL := IsRightToLeft;
  if (uffBiDi in fmt^.UsedFormattingFields) then
    case Worksheet.ReadBiDiMode(lCell) of
      bdRTL : RTL := true;
      bdLTR : RTL := false;
    end;

  // Text rotation
  if (uffTextRotation in fmt^.UsedFormattingFields)
    then txtRot := fmt^.TextRotation
    else txtRot := trHorizontal;

  // vertical alignment
  if (uffVertAlign in fmt^.UsedFormattingFields)
    then vertAlign := fmt^.VertAlignment
    else vertAlign := vaDefault;
  if vertAlign = vaDefault then
    vertAlign := vaBottom;

  // Horizontal alignment
  if (uffHorAlign in fmt^.UsedFormattingFields)
    then horAlign := fmt^.HorAlignment
    else horAlign := haDefault;
  if (horAlign = haDefault) then
  begin
    if (lCell^.ContentType in [cctNumber, cctDateTime]) then
    begin
      if RTL then
        horAlign := haLeft
      else
        horAlign := haRight;
    end else
    if (lCell^.ContentType in [cctBool]) then
      horAlign := haCenter
    else begin
      if RTL then
        horAlign := haRight
      else
        horAlign := haLeft;
    end;
  end;

  // Font index
  if (uffFont in fmt^.UsedFormattingFields)
    then fntIndex := fmt^.FontIndex
    else fntIndex := DEFAULT_FONTINDEX;

  // Font color as derived from number format
  numFmtColor := clNone;
  if not IsNaN(lCell^.NumberValue) and (uffNumberFormat in fmt^.UsedFormattingFields) then
  begin
    numFmt := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);
    if numFmt <> nil then
    begin
      sidx := 0;
      if (Length(numFmt.Sections) > 1) and (lCell^.NumberValue < 0) then
        sidx := 1
      else
      if (Length(numFmt.Sections) > 2) and (lCell^.NumberValue = 0) then
        sidx := 2;
      if (nfkHasColor in numFmt.Sections[sidx].Kind) then
        numFmtColor := numFmt.Sections[sidx].Color and $00FFFFFF;
    end;
  end;

  InflateRect(ARect, -constCellPadding, -constCellPadding);

  InternalDrawTextInCell(txt, ARect, horAlign, vertAlign, txtRot, wrapped,
    fntIndex, numfmtColor, lCell^.RichTextParams, RTL);
end;

{@@ ----------------------------------------------------------------------------
  This procedure is called when editing of a cell is completed. It determines
  the worksheet cell and writes the text into the worksheet. Tries to keep the
  format of the cell, but if it is a new cell, or the content type has changed,
  tries to figure out the content type (number, date/time, text).
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.EditingDone;
var
  //oldText: String;
  cell: PCell;
begin
  if (not EditorShowing) and FEditing then
  begin
    {
    oldText := GetCellText(Col, Row);
    if oldText <> FEditText then
    }
    if FOldEditorText <> FEditText then
    begin
      cell := Worksheet.GetCell(GetWorksheetRow(Row), GetWorksheetCol(Col));
      if Worksheet.IsMerged(cell) then
        cell := Worksheet.FindMergeBase(cell);
      if (FEditText <> '') and (FEditText[1] = '=') then
        Worksheet.WriteFormula(cell, Copy(FEditText, 2, Length(FEditText)), true)
      else
        Worksheet.WriteCellValueAsString(cell, FEditText);
      FEditText := '';
      FOldEditorText := '';
    end;
    inherited EditingDone;
  end;
  FEditing := false;
  FEnhEditMode := false;
end;

function TsCustomWorksheetGrid.EditorByStyle(Style: TColumnButtonStyle): TWinControl;
begin
  if (Style = cbsAuto) and (FLineMode = elmMultiLine) then
    Result := FMultilineStringEditor
  else
    Result := inherited;
end;

{@@ ----------------------------------------------------------------------------
  The BeginUpdate/EndUpdate pair suppresses unnecessary painting of the grid.
  Call BeginUpdate to stop refreshing the grid, and call EndUpdate to release
  the lock and to repaint the grid again.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.EndUpdate(ARefresh: Boolean = true);
begin
  inherited EndUpdate(false);
  dec(FLockCount);
  if (FLockCount = 0) and ARefresh then
    VisualChange;
end;

{@@ ----------------------------------------------------------------------------
  Executes a hyperlink stored in the FHyperlinkCell
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.ExecuteHyperlink;
var
  hlink: TsHyperlink;
  target, bookmark: String;
  sheetname: String;
  sheet: TsWorksheet;
  r, c: Cardinal;
begin
  if FHyperlinkCell = nil then
    exit;

  hlink := Worksheet.ReadHyperlink(FHyperlinkCell);
  SplitHyperlink(hlink.Target, target, bookmark);
  if target = '' then begin
    // Goes to a cell within the current workbook
    if ParseSheetCellString(bookmark, sheetname, r, c) then
    begin
      if sheetname <> '' then
      begin
        sheet := Workbook.GetWorksheetByName(sheetname);
        if sheet = nil then
          raise Exception.CreateFmt(rsWorksheetNotFound, [sheetname]);
        Workbook.SelectWorksheet(sheet);
      end;
      Worksheet.SelectCell(r, c);
    end else
      raise Exception.CreateFmt(rsNoValidHyperlinkInternal, [hlink.Target]);
  end else
    // Fires the OnClickHyperlink event which should open a file or a URL
    if Assigned(FOnClickHyperlink) then FOnClickHyperlink(self, hlink);
end;

{@@ ----------------------------------------------------------------------------
  Copies the borders of a cell to the correspondig edges of its neighbors.
  This avoids the nightmare of changing borders due to border conflicts
  of adjacent cells.

  @param  ACell  Pointer to the cell
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.FixNeighborCellBorders(ACell: PCell);

  procedure SetNeighborBorder(NewRow, NewCol: Cardinal;
    ANewBorder: TsCellBorder; const ANewBorderStyle: TsCellBorderStyle;
    AInclude: Boolean);
  var
    neighbor: PCell;
    border: TsCellBorders;
  begin
    neighbor := Worksheet.FindCell(NewRow, NewCol);
    if neighbor <> nil then
    begin
      border := Worksheet.ReadCelLBorders(neighbor);
      if AInclude then
      begin
        Include(border, ANewBorder);
        Worksheet.WriteBorderStyle(NewRow, NewCol, ANewBorder, ANewBorderStyle);
      end else
        Exclude(border, ANewBorder);
      Worksheet.WriteBorders(NewRow, NewCol, border);
    end;
  end;

var
  fmt: PsCellFormat;
begin
  if (Worksheet = nil) or (ACell = nil) then
    exit;

  fmt := Worksheet.GetPointerToEffectiveCellFormat(ACell);
  with ACell^ do
  begin
//    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if Col > 0 then
      SetNeighborBorder(Row, Col-1, cbEast, fmt^.BorderStyles[cbWest], cbWest in fmt^.Border);
    SetNeighborBorder(Row, Col+1, cbWest, fmt^.BorderStyles[cbEast], cbEast in fmt^.Border);
    if Row > 0 then
      SetNeighborBorder(Row-1, Col, cbSouth, fmt^.BorderStyles[cbNorth], cbNorth in fmt^.Border);
    SetNeighborBorder(Row+1, Col, cbNorth, fmt^.BorderStyles[cbSouth], cbSouth in fmt^.Border);
  end;
end;

(*
{@@ ----------------------------------------------------------------------------
  The "colors" used by the spreadsheet are indexes into the workbook's color
  palette. If the user wants to set a color to a particular RGB value this is
  not possible in general. The method FindNearestPaletteIndex finds the bast
  matching color in the palette.

  @param  AColor  Color index into the workbook's palette
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.FindNearestPaletteIndex(AColor: TColor): TsColor;
begin
  Result := fpsVisualUtils.FindNearestPaletteIndex(Workbook, AColor);
end;
  *)
                                          (*
{@@ ----------------------------------------------------------------------------
  Notification by the workbook link that a cell has been modified. --> Repaint.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.CellChanged(ACell: PCell);
begin
  Unused(ACell);
  Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Notification by the workbook link that another cell has been selected
  --> select the cell in the grid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.CellSelected(ASheetRow, ASheetCol: Cardinal);
var
  grow, gcol: Integer;
begin
  if (FWorkbookLink <> nil) then
  begin
    grow := GetGridRow(ASheetRow);
    gcol := GetGridCol(ASheetCol);
    if (grow <> Row) or (gcol <> Col) then begin
      Row := grow;
      Col := gcol;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Notification by the workbook link that a new workbook is available.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.WorkbookChanged;
begin
  WorksheetChanged;
end;

{@@ ----------------------------------------------------------------------------
  Notification by the workbook link that a new worksheet has been selected from
  the current workbook. Is also called internally from WorkbookChanged.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.WorksheetChanged;
begin
  if Worksheet <> nil then begin
    //Worksheet.OnChangeCell := @ChangedCellHandler;
    //Worksheet.OnChangeFont := @ChangedFontHandler;
    ShowHeaders := (soShowHeaders in Worksheet.Options);
    ShowGridLines := (soShowGridLines in Worksheet.Options);
    if (soHasFrozenPanes in Worksheet.Options) then begin
      FrozenCols := Worksheet.LeftPaneWidth;
      FrozenRows := Worksheet.TopPaneHeight;
    end else begin
      FrozenCols := 0;
      FrozenRows := 0;
    end;
    Row := FrozenRows;
    Col := FrozenCols;
  end;
  Setup;
end;
                                            *)

{@@ ----------------------------------------------------------------------------
  Returns the background color of a cell. The color is given as an index into
  the workbook's color palette.

  @param  ACol  Grid column index of the cell
  @param  ARow  Grid row index of the cell
  @result Color index of the cell's background color.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetBackgroundColor(ACol, ARow: Integer): TsColor;
var
  cell: PCell;
begin
  Result := scNotDefined;
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Result := Worksheet.ReadBackgroundColor(cell);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the background color of a cell range defined by a rectangle. The color
  is given as an index into the workbook's color palette. If the colors are
  different from cell to cell the value scUndefined is returned.

  @param  ALeft   Index of the left column of the cell range
  @param  ATop    Index of the top row of the cell range
  @param  ARight  Index of the right column of the cell range
  @param  ABottom Index of the bottom row of the cell range
  @return Color index common to all cells within the selection. If the cells'
          background colors are different the value scUndefined is returned.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetBackgroundColors(ALeft, ATop, ARight, ABottom: Integer): TsColor;
var
  c, r: Integer;
  clr: TsColor;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetBackgroundColor(ALeft, ATop);
  clr := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do
    begin
      Result := GetBackgroundColor(c, r);
      if Result <> clr then
      begin
        Result := scNotDefined;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellBiDiMode(ACol,ARow: Integer): TsBiDiMode;
var
  cell: PCell;
begin
  cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
  if cell <> nil then
    Result := Worksheet.ReadBiDiMode(cell) else
    Result := bdDefault;
end;

{@@ ----------------------------------------------------------------------------
  Returns the cell borders which are drawn around a given cell.
  If the cell is part of a merged block then the borders of the merge base
  are applied to the location of the cell (no inner borders for merged cells).

  @param  ACol  Grid column index of the cell
  @param  ARow  Grid row index of the cell
  @return Set with flags indicating where borders are drawn (top/left/right/bottom)
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorder(ACol, ARow: Integer): TsCellBorders;
var
  cell: PCell;
  base: PCell;
  r, c, r1, c1, r2, c2: Cardinal;
begin
  Result := [];
  if Assigned(Worksheet) then
  begin
    r := GetWorksheetRow(ARow);
    c := GetWorksheetCol(ACol);
    cell := Worksheet.FindCell(r, c);
    if Worksheet.IsMerged(cell) then
    begin
      Worksheet.FindMergedRange(cell, r1, c1, r2, c2);
      base := Worksheet.FindCell(r1, c1);
      Result := Worksheet.ReadCellBorders(base);
      if (cbNorth in Result) and (r > r1) then Exclude(Result, cbNorth);
      if (cbSouth in Result) and (r < r2) then Exclude(Result, cbSouth);
      if IsRightToLeft then
      begin
        if (cbEast in Result) and (c > c1) then Exclude(Result, cbEast);
        if (cbWest in Result) and (c < c2) then Exclude(Result, cbWest);
      end else
      begin
        if (cbWest in Result) and (c > c1) then Exclude(Result, cbWest);
        if (cbEast in Result) and (c < c2) then Exclude(Result, cbEast);
      end;
    end else
      Result := Worksheet.ReadCellBorders(cell);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the cell borders which are drawn around a given rectangular cell range.

  @param  ALeft   Index of the left column of the cell range
  @param  ATop    Index of the top row of the cell range
  @param  ARight  Index of the right column of the cell range
  @param  ABottom Index of the bottom row of the cell range
  @return Set with flags indicating where borders are drawn (top/left/right/bottom)
          If the individual cells within the range have different borders an
          empty set is returned.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorders(ALeft, ATop, ARight, ABottom: Integer): TsCellBorders;
var
  c, r: Integer;
  b: TsCellBorders;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetCellBorder(ALeft, ATop);
  b := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do
    begin
      Result := GetCellBorder(c, r);
      if Result <> b then
      begin
        Result := [];
        exit;
      end;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the style of the cell border line drawn along the edge specified
  by the parameter ABorder of a cell. The style is defined by line style and
  line color.

  If the cell belongs to a merged block then the border styles of the merge
  base are returned.

  @param   ACol     Grid column index of the cell
  @param   ARow     Grid row index of the cell
  @param   ABorder  Identifier of the border at which the line will be drawn
                    (see TsCellBorder)
  @return  CellBorderStyle record containing information on line style and
           line color.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorderStyle(ACol, ARow: Integer;
  ABorder: TsCellBorder): TsCellBorderStyle;
var
  cell: PCell;
begin
  Result := DEFAULT_BORDERSTYLES[ABorder];
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if Worksheet.IsMerged(cell) then
      cell := Worksheet.FindMergeBase(cell);
    Result := Worksheet.ReadCellBorderStyle(cell, ABorder);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the style of the cell border line drawn along the edge specified
  by the parameter ABorder of a range of cells defined by the rectangle of
  column and row indexes. The style is defined by linestyle and line color.

  @param   ALeft   Index of the left column of the cell range
  @param   ATop    Index of the top row of the cell range
  @param   ARight  Index of the right column of the cell range
  @param   ABottom Index of the bottom row of the cell range
  @param   ABorder  Identifier of the border where the line will be drawn
                    (see TsCellBorder)
  @return  CellBorderStyle record containing information on line style and
           line color.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellBorderStyles(ALeft, ATop, ARight, ABottom: Integer;
  ABorder: TsCellBorder): TsCellBorderStyle;
var
  c, r: Integer;
  bs: TsCellBorderStyle;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetCellBorderStyle(ALeft, ATop, ABorder);
  bs := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do
    begin
      Result := GetCellBorderStyle(c, r, ABorder);
      if (Result.LineStyle <> bs.LineStyle) or (Result.Color <> bs.Color) then
      begin
        Result := DEFAULT_BORDERSTYLES[ABorder];
        exit;
      end;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the comment assigned to a cell.

  @param   ACol     Grid column index of the cell
  @param   ARow     Grid row index of the cell
  @return  String used as a cell comment.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellComment(ACol, ARow: Integer): String;
begin
  if Worksheet <> nil then
    Result := Worksheet.ReadComment(GetWorksheetRow(ARow), GetWorksheetCol(ACol))
  else
    Result :='';
end;

{@@ ----------------------------------------------------------------------------
  Returns the (LCL) font to be used when painting text in a cell.

  @param   ACol     Grid column index of the cell
  @param   ARow     Grid row index of the cell
  @return  Font usable when painting on a canvas.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellFont(ACol, ARow: Integer): TFont;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := nil;
  if (Workbook <> nil) then
  begin
    fnt := Workbook.GetDefaultFont;
    if  (Worksheet <> nil) then
    begin
      cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
      fnt := Worksheet.ReadCellFont(cell);
    end;
    Convert_sFont_to_Font(fnt, FCellFont);
    Result := FCellFont;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the (LCL) font to be used when painting text in the cells defined
  by the rectangle of row/column indexes.

  @param   ALeft   Index of the left column of the cell range
  @param   ATop    Index of the top row of the cell range
  @param   ARight  Index of the right column of the cell range
  @param   ABottom Index of the bottom row of the cell range
  @return  Font usable when painting on a canvas.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellFonts(ALeft, ATop, ARight, ABottom: Integer): TFont;
var
  r1,c1,r2,c2: Cardinal;
  sFont, sDefFont: TsFont;
  cell: PCell;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetCellFont(ALeft, ATop);
  sDefFont := Workbook.GetDefaultFont;  // Default font
  r1 := GetWorksheetRow(ATop);
  c1 := GetWorksheetCol(ALeft);
  r2 := GetWorksheetRow(ABottom);
  c2 := GetWorksheetRow(ARight);
  for cell in Worksheet.Cells.GetRangeEnumerator(r1, c1, r2, c2) do
  begin
    sFont := Worksheet.ReadCellFont(cell);
    if (sFont.FontName <> sDefFont.FontName) and (sFont.Size <> sDefFont.Size)
      and (sFont.Style <> sDefFont.Style) and (sFont.Color <> sDefFont.Color)
    then
    begin
      Convert_sFont_to_Font(sDefFont, FCellFont);
      Result := FCellFont;
      exit;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the height (in pixels) of the cell at ACol/ARow (of the grid).

  @param   ACol  Grid column index of the cell
  @param   ARow  Grid row index of the cell
  @result  Height of the cell in pixels. Wrapped text is handled correctly.
           Value contains the zoom factor.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellHeight(ACol, ARow: Integer): Integer;
var
  lCell: PCell;
  s: String;
  wrapped: Boolean;
  cellR: TRect;
  r1,c1,r2,c2: Cardinal;
  fmt: PsCellFormat;
  fntIndex: Integer;
  txtRot: TsTextRotation;
  RTL: Boolean;
begin
  Result := 0;
  if (ACol < FHeaderCount) or (ARow < FHeaderCount) then
    exit;
  if Worksheet = nil then
    exit;

  lCell := Worksheet.FindCell(ARow-FHeaderCount, ACol-FHeaderCount);
  if lCell <> nil then
  begin
    cellR := CellRect(ACol, ARow);
    if Worksheet.IsMerged(lCell) then
    begin
      Worksheet.FindMergedRange(lCell, r1, c1, r2, c2);
      if r1 <> r2 then
        // If the merged range encloses several rows we skip automatic row height
        // determination since only the height of the first row of the block
        // (containing the merge base cell) would change which is very confusing.
        exit;
      cellR := CellRect(LongInt(c1)+FHeaderCount, ARow);
      cellR.Right := CellRect(LongInt(c2)+FHeaderCount, ARow).Right;
    end;
    InflateRect(cellR, -constCellPadding, -constCellPadding);

    s := GetCellText(ACol, ARow, false);
    if s = '' then
      exit;

    DoPrepareCanvas(ACol, ARow, []);

//    fmt := Workbook.GetPointerToCellFormat(lCell^.FormatIndex);
    fmt := Worksheet.GetPointerToEffectiveCellFormat(lCell);
    if (uffFont in fmt^.UsedFormattingFields) then
      fntIndex := fmt^.FontIndex else fntIndex := DEFAULT_FONTINDEX;
    if (uffTextRotation in fmt^.UsedFormattingFields) then
      txtRot := fmt^.TextRotation else txtRot := trHorizontal;
    wrapped := (uffWordWrap in fmt^.UsedFormattingFields);

    RTL := IsRightToLeft;
    if (uffBiDi in fmt^.UsedFormattingFields) then
      case fmt^.BiDiMode of
        bdRTL: RTL := true;
        bdLTR: RTL := false;
      end;

    case txtRot of
      trHorizontal: ;
      rt90DegreeClockwiseRotation,
      rt90DegreeCounterClockwiseRotation:
        cellR := Rect(0, 0, MaxInt, MaxInt);
      rtStacked:
        cellR := Rect(0, 0, MaxInt, MaxInt);
    end;

    Result := RichTextHeight(Canvas, Workbook, cellR, s, lCell^.RichTextParams,
                fntIndex, txtRot, wrapped, RTL, ZoomFactor)
            + 2 * constCellPadding;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This function defines the text to be displayed as a cell hint. By default, it
  is the comment and/or the hyperlink attached to a cell; it can further be
  modified by using the OnGetCellHint event.
  Option goCellHints must be active for the cell hint feature to work.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellHintText(ACol, ARow: Integer): String;
var
  cell: PCell;
  hlink: PsHyperlink;
  comment: String;
begin
  Result := '';

  if Worksheet = nil then
    exit;

  cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
  if cell = nil then
    exit;

  // Read comment
  comment := Worksheet.ReadComment(cell);

  // Read hyperlink info
  if Worksheet.HasHyperlink(cell) then begin
    hlink := Worksheet.FindHyperlink(cell);
    if hlink <> nil then
    begin
      if hlink^.ToolTip <> '' then
        Result := hlink^.ToolTip
      else
        Result := Format('Hyperlink: %s' + LineEnding + rsStdHyperlinkTooltip,
          [hlink^.Target]
        );
    end;
  end;

  // Combine comment and hyperlink
  if (Result <> '') and (comment <> '') then
    Result := comment + LineEnding + LineEnding + Result
  else
  if (Result = '') and (comment <> '') then
    Result := comment;

  // Call hint event handler
  if Assigned(OnGetCellHint) then
    OnGetCellHint(self, ACol, ARow, Result);
end;

function TsCustomWorksheetGrid.GetCells(ACol, ARow: Integer): String;
var
  msg: TGridMessage;
begin
  if (Editor <> nil) and Editor.Visible then
  begin
    msg.LclMsg.msg := GM_GETVALUE;
    msg.Grid := Self;
    msg.Col := ACol;
    msg.Row := ARow;
    msg.Value := ''; //GetCells(FCol, FRow);
    Editor.Dispatch(msg);
    Result := msg.value;
  end else
    Result := GetCellText(ACol, ARow);
end;

{@@ ----------------------------------------------------------------------------
  This function returns the text to be shown in a grid cell. The text is looked
  up in the corresponding cell of the worksheet by calling its ReadAsUTF8Text
  method. In case of "stacked" text rotation, line endings are inserted after
  each character.

  @param   ACol   Grid column index of the cell
  @param   ARow   Grid row index of the cell
  @param   ATrim  If true show replacement characters if numerical data
                  are wider than cell.
  @return  Text to be displayed in the cell.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetCellText(ACol, ARow: Integer;
  ATrim: Boolean = true): String;
var
  cell: PCell;
  r, c: Integer;
begin
  Result := '';

  if ShowHeaders then
  begin
    // Headers
    if (ARow = 0) and (ACol = 0) then
      exit;
    if (ARow = 0) then
    begin
      Result := GetColString(ACol - FHeaderCount);
      exit;
    end
    else
    if (ACol = 0) then
    begin
      Result := IntToStr(ARow);
      exit;
    end;
  end;

  if Worksheet <> nil then
  begin
    r := ARow - FHeaderCount;
    c := ACol - FHeaderCount;
    cell := Worksheet.FindCell(r, c);
    if cell <> nil then
    begin
      if ATrim then
        Result := TrimToCell(cell)
      else
        Result := Worksheet.ReadAsText(cell);
      {
      if Worksheet.ReadTextRotation(cell) = rtStacked then
      begin
        s := Result;
        Result := '';
        for i:=1 to Length(s) do
        begin
          Result := Result + s[i];
          if i < Length(s) then Result := Result + LineEnding;
        end;
      end;
      }
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the width of the fixed header column 0, in pixels
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetDefaultHeaderColWidth: Integer;
begin
  Result := Canvas.TextWidth(' 9999999 ');
end;

{@@ ----------------------------------------------------------------------------
  Determines the text to be passed to the cell editor. The text is determined
  from the underlying worksheet cell, but it is possible to intercept this by
  adding a handler for the OnGetEditText event.

  @param   ACol   Grid column index of the cell being edited
  @param   ARow   Grid row index of the grid cell being edited
  @return  Text to be passed to the cell editor.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetEditText(ACol, ARow: Integer): string;
var
  cell: PCell;
begin
  if FEnhEditMode then   // Initiated by "F2"
  begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if Worksheet.IsMerged(cell) then
      cell := Worksheet.FindMergeBase(cell);
    Result := Worksheet.ReadFormulaAsString(cell, true);
    if Result <> '' then
    begin
      if Result[1] <> '=' then Result := '=' + Result;
    end else
    if cell <> nil then
      case cell^.ContentType of
        cctNumber:
          Result := FloatToStr(cell^.NumberValue);
        cctDateTime:
          if cell^.DateTimeValue < 1.0 then
            Result := FormatDateTime('tt', cell^.DateTimeValue)
          else
            Result := FormatDateTime('c', cell^.DateTimeValue);
        else
          Result := Worksheet.ReadAsText(cell);
      end
    else
      Result := '';
  end else
    Result := GetCellText(aCol, aRow);
  if Assigned(OnGetEditText) then
    OnGetEditText(Self, aCol, aRow, Result);
end;

{@@ ----------------------------------------------------------------------------
  Determines the style of the border between a cell and its neighbor given by
  ADeltaCol and ADeltaRow (one of them must be 0, the other one can only be +/-1).
  ACol and ARow are in grid units.
  Result is FALSE if there is no border line.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetBorderStyle(ACol, ARow, ADeltaCol, ADeltaRow: Integer;
  ACell: PCell; out ABorderStyle: TsCellBorderStyle): Boolean;
var
  neighborcell: PCell;
  border, neighborborder: TsCellBorder;
  r, c: Cardinal;
begin
  Result := true;

  if (ADeltaCol = -1) and (ADeltaRow = 0) then
  begin
    border := cbWest;
    neighborborder := cbEast;
  end else
  if (ADeltaCol = +1) and (ADeltaRow = 0) then
  begin
    border := cbEast;
    neighborborder := cbWest;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = -1) then
  begin
    border := cbNorth;
    neighborborder := cbSouth;
  end else
  if (ADeltaCol = 0) and (ADeltaRow = +1) then
  begin
    border := cbSouth;
    neighborBorder := cbNorth;
  end else
    raise Exception.Create('[TsCustomWorksheetGrid] Incorrect col/row for GetBorderStyle.');

  r := GetWorksheetRow(ARow);
  c := GetWorksheetCol(ACol);
  if (longint(r) + ADeltaRow < 0) or (longint(c) + ADeltaCol < 0) then
    neighborcell := nil
  else
    neighborcell := Worksheet.FindCell(longint(r) + ADeltaRow, longint(c) + ADeltaCol);

  // Only cell has border, but neighbor has not
  if HasBorder(ACell, border) and not HasBorder(neighborCell, neighborBorder) then
  begin
    if Worksheet.InSameMergedRange(ACell, neighborcell) then
      result := false
    else
      ABorderStyle := GetCellBorderStyle(ACol, ARow, border);
  end
  else
  // Only neighbor has border, cell has not
  if not HasBorder(ACell, border) and HasBorder(neighborCell, neighborBorder) then
  begin
    if Worksheet.InSameMergedRange(ACell, neighborcell) then
      result := false
    else
      ABorderStyle := GetCellBorderStyle(ACol+ADeltaCol, ARow+ADeltaRow, neighborborder);
  end
  else
  // Both cells have shared border -> use top or left border
  if HasBorder(ACell, border) and HasBorder(neighborCell, neighborBorder) then
  begin
    if Worksheet.InSameMergedRange(ACell, neighborcell) then
      result := false
    else
    if (border in [cbNorth, cbWest]) then
      ABorderStyle := GetCellBorderStyle(ACol+ADeltaCol, ARow+ADeltaRow, neighborborder)
    else
      ABorderStyle := GetCellBorderStyle(ACol, ARow, border);
  end else
    Result := false;
end;

{@@ ----------------------------------------------------------------------------
  Converts a column index of the worksheet to a column index usable in the grid.
  This is required because worksheet indexes always start at zero while
  grid indexes also have to account for the column/row headers.

  @param  ASheetCol   Worksheet column index
  @return Grid column index
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetGridCol(ASheetCol: Cardinal): Integer;
begin
  Result := Integer(ASheetCol) + FHeaderCount
end;

{@@ ----------------------------------------------------------------------------
  Converts a row index of the worksheet to a row index usable in the grid.
  This is required because worksheet indexes always start at zero while
  grid indexes also have to account for the column/row headers.

  @param  ASheetRow   Worksheet row index
  @return Grid row index
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetGridRow(ASheetRow: Cardinal): Integer;
begin
  Result := Integer(ASheetRow) + FHeaderCount;
end;

{@@ ----------------------------------------------------------------------------
  Returns a list of worksheets contained in the file. Useful for assigning to
  user controls like TabControl, Combobox etc. in order to select a sheet.

  @param  ASheets  List of strings containing the names of the worksheets of
                   the workbook
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.GetSheets(const ASheets: TStrings);
var
  i: Integer;
begin
  ASheets.Clear;
  if Assigned(Workbook) then
    for i:=0 to Workbook.GetWorksheetCount-1 do
      ASheets.Add(Workbook.GetWorksheetByIndex(i).Name);
end;

{@@ ----------------------------------------------------------------------------
  Calculates the index of the worksheet column that is displayed in the
  given column of the grid. If the sheet headers are turned on, both numbers
  differ by 1, otherwise they are equal. Saves an "if" in cases.

  @param   AGridCol   Index of a grid column
  @return  Index of a the corresponding worksheet column
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetWorksheetCol(AGridCol: Integer): cardinal;
begin
  if (FHeaderCount > 0) and (AGridCol = 0) then
    Result := UNASSIGNED_ROW_COL_INDEX
  else
    Result := AGridCol - FHeaderCount;
end;

{@@ ----------------------------------------------------------------------------
  Calculates the index of the worksheet row that is displayed in the
  given row of the grid. If the sheet headers are turned on, both numbers
  differ by 1, otherwise they are equal. Saves an "if" in some cases.

  @param    AGridRow  Index of a grid row
  @resturn  Index of the corresponding worksheet row.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.GetWorksheetRow(AGridRow: Integer): Cardinal;
begin
  if (FHeaderCount > 0) and (AGridRow = 0) then
    Result := UNASSIGNED_ROW_COL_INDEX
  else
    Result := AGridRow - FHeaderCount;
end;

{@@ ----------------------------------------------------------------------------
  Returns true if the cell has the given border. In case of merged cell the
  borders of the merge base are checked. Inner merged cells don't have a border.

  @param  ACell    Pointer to cell considered
  @param  ABorder  Indicator for border to be checked for visibility
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.HasBorder(ACell: PCell; ABorder: TsCellBorder): Boolean;
var
  base: PCell;
  r1, c1, r2, c2: Cardinal;
begin
  if Worksheet = nil then
    result := false
  else
  if Worksheet.IsMerged(ACell) then
  begin
    Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    base := Worksheet.FindCell(r1, c1);
    Result := ABorder in Worksheet.ReadCellBorders(base);
    case ABorder of
      cbNorth : if ACell^.Row > r1 then Result := false;
      cbSouth : if ACell^.Row < r2 then Result := false;
      cbEast  : if ACell^.Col < c2 then Result := false;
      cbWest  : if ACell^.Col > c1 then Result := false;
      {
      cbEast  : if (IsRightToLeft and (ACell^.Col > c1)) or
                   (not IsRightToLeft and (ACell^.Col < c2)) then Result := false;
      cbWest  : if (IsRightToLeft and (ACell^.Col < c2)) or
                   (not IsRightToLeft and (ACell^.Col > c1)) then Result := false;
                   }
    end;
  end else
    Result := ABorder in Worksheet.ReadCellBorders(ACell);
end;

{@@ ----------------------------------------------------------------------------
  HeaderSizing is called while a column width or row height is resized by the
  mouse. Is overridden here to enforce a grid repaint if merged cells are
  affected by the resizing column/row. Otherwise parts of the merged cells would
  not be updated if the cell text moves during the resizing action.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.HeaderSizing(const IsColumn:boolean;
  const AIndex,ASize:Integer);
var
  gc, gr: Integer;
  sr1, sr2, sc1, sc2, si: Cardinal;
  cell: PCell;
begin
  inherited;

  if Worksheet = nil then
    exit;

  // replaint the grid if merged cells are affected by the resizing col/row.
  si := IfThen(IsColumn, GetWorksheetCol(AIndex), GetWorksheetRow(AIndex));
  for gc := GetFirstVisibleColumn to GetLastVisibleColumn do
  begin
    for gr := GetFirstVisibleRow to GetLastVisibleRow do
    begin
      cell := Worksheet.FindCell(gr, gc);
      if Worksheet.IsMerged(cell) then begin
        Worksheet.FindMergedRange(cell, sr1, sc1, sr2, sc2);
        if IsColumn and InRange(si, sc1, sc2) or
           (not IsColumn) and InRange(si, sr1, sr2) then
        begin
          InvalidateGrid;
          exit;
        end;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Inherited from TCustomGrid. Is called when column widths or row heights
  have changed. Stores the new column width or row height in the worksheet.

  @param   IsColumn   Specifies whether the changed parameter is a column width
                      (true) or a row height (false)
  @param   Index      Index of the changed column or row
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.HeaderSized(IsColumn: Boolean; AIndex: Integer);
const
  EPS = 0.1;
var
  idx: Integer;
  w, h, wdef, hdef: Single;
begin
  if (Worksheet = nil) or (FZoomLock <> 0) then
    exit;

  if IsColumn then
  begin
    w := CalcWorksheetColWidth(ColWidths[AIndex]);   // w and wdef are at 100% zoom
    wdef := Worksheet.ReadDefaultColWidth(Workbook.Units);
    if (AIndex >= FHeaderCount) and not SameValue(w, wdef, EPS) then begin
      idx := GetWorksheetCol(AIndex);
      Worksheet.WriteColWidth(idx, w, Workbook.Units);
    end;
  end else
  begin
    h := CalcWorksheetRowHeight(RowHeights[AIndex]);
    hdef := Worksheet.ReadDefaultRowHeight(Workbook.Units);
    if (AIndex >= FHeaderCount) and not SameValue(h, hdef, EPS) then begin
      idx := GetWorksheetRow(AIndex);
      Worksheet.WriteRowHeight(idx, h, Workbook.Units);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Clicking into cells with hyperlinks poses a user-interface problem:
  normally the cell should go into edit mode. But with hyperlinks a click should
  also execute the hyperlink. How to distinguish both cases?
  In order to keep both features for hyperlinks we follow a strategy similar to
  Excel: a short click selects the cell for editing as usual; a longer click
  opens the hyperlink by means of a timer ("FHyperlinkTimer") (in Excel, in
  fact, the behavior is opposite, but this one here is easier to implement.)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.HyperlinkTimerElapsed(Sender: TObject);
begin
  if FHyperlinkTimer.Enabled then begin
    FHyperlinkTimer.Enabled := false;
    FGridState := gsNormal;  // this prevents selecting a cell block
    EditorMode := false;     // this prevents editing the clicked cell
    ExecuteHyperlink;        // Execute the hyperlink
    FHyperlinkCell := nil;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Inserts an empty column before the column specified
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.InsertCol(AGridCol: Integer);
var
  c: Cardinal;
begin
  if AGridCol < FHeaderCount then
    exit;

//  if LongInt(Worksheet.GetLastColIndex) + 1 + FHeaderCount >= ColCount then //FInitColCount then
  if GetGridCol(Worksheet.GetLastColIndex + 1) >= ColCount then
    ColCount := ColCount + 1;
  c := GetWorksheetCol(AGridCol);
//  c := AGridCol - FHeaderCount;
  Worksheet.InsertCol(c);

  UpdateColWidths(AGridCol);
end;

{@@ ----------------------------------------------------------------------------
  Inserts an empty row before the row specified
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.InsertRow(AGridRow: Integer);
var
  r: Cardinal;
begin
  if AGridRow < FHeaderCount then
    exit;

//  if LongInt(Worksheet.GetlastRowIndex) + 1 + FHeaderCount >= FInitRowCount then
  if GetGridRow(Worksheet.GetLastRowIndex + 1) >= RowCount then
    RowCount := RowCount + 1;
  r := GetWorksheetRow(AGridRow);
  Worksheet.InsertRow(r);

  // Calculate row height if new row
  UpdateRowHeight(AGridRow, true);

  // Update following row heights because their index has changed.
  UpdateRowHeights(AGridRow);
end;

procedure TsCustomWorksheetGrid.InternalDrawCell(ACol, ARow: Integer;
  AClipRect, ACellRect: TRect; AState: TGridDrawState);

  function IsPushCellActive: boolean;
  begin
    with GCache do
      Result := (PushedCell.X <> -1) and (PushedCell.Y <> -1);
  end;

var
  rgn: HRGN;
begin
  with GCache do begin
    if (ACol = HotCell.x) and (ARow = HotCell.y) and not IsPushCellActive() then begin
      Include(AState, gdHot);
      HotCellPainted := True;
    end;
    if ClickCellPushed and (ACol = PushedCell.x) and (ARow = PushedCell.y) then begin
      Include(AState, gdPushed);
    end;
  end;
  Canvas.SaveHandleState;
  try
    rgn := CreateRectRgn(AClipRect.Left, AClipRect.Top, AClipRect.Right, AClipRect.Bottom);
    SelectClipRgn(Canvas.Handle, rgn);
    DrawCell(ACol, ARow, ACellRect, AState);
    DeleteObject(rgn);
  finally
    Canvas.RestoreHandleState;
  end;
end;

{ Draws the cells in the specified row. Drawing takes care of text overflow
  and merged cells.
  AClipRect covers the paintable row, painting outside will be clipped. }
procedure TsCustomWorksheetGrid.InternalDrawRow(ARow, AFirstCol, ALastCol: Integer;
  AClipRect: TRect);
var
  sr: Cardinal;
  scLastUsed: Cardinal;
  sr1, sc1, sr2, sc2: Cardinal;
  gr, gc, gc1, gc2, gcNext, gcLast, gcLastUsed: Integer;
  i: Integer;
  tmp: Integer = 0;
  cell: PCell;
  fmt: PsCellFormat;
  rct, clip_rct, commentcell_rct, temp_rct: TRect;
  gds: TGridDrawState;
  clipArea: TRect;

begin
  sr := GetWorksheetRow(ARow);
  scLastused := Worksheet.GetLastColIndex;
  gc := AFirstCol;
  gcLast := ALastCol;

  clipArea := Canvas.ClipRect;

  with GCache.VisibleGrid do
  begin
    // Because of possible cell overflow from cells left of the visible range
    // we have to seek to the left for the first occupied text cell
    // and start painting from here.
    if FTextOverflow and (sr <> UNASSIGNED_ROW_COL_INDEX) and Assigned(Worksheet) then
      while (gc > FHeaderCount) do
      begin
        dec(gc);
        cell := Worksheet.FindCell(sr, GetWorksheetCol(gc));
        // Empty cell --> proceed with next cell to the left
        if (cell = nil) or (cell^.ContentType = cctEmpty) or
           ((cell^.ContentType = cctUTF8String) and (cell^.UTF8StringValue = ''))
        then
          Continue;
        // Overflow possible from non-merged, non-right-aligned, horizontal label cells
        fmt := Worksheet.GetPointerToEffectiveCellFormat(cell);
        if (not Worksheet.IsMerged(cell)) and
           (cell^.ContentType = cctUTF8String) and
           not (uffTextRotation in fmt^.UsedFormattingFields) and
           (uffHorAlign in fmt^.UsedFormattingFields) and (fmt^.HorAlignment <> haRight)
        then
          Break;
        // All other cases --> no overflow --> return to initial left cell
        gc := AFirstCol;
        break;
      end;

    // Now find the last column. Again text can overflow into the visible area
    // from invisible cells at the right.
    gcLast := ALastCol;
    if FTextOverflow and (sr <> UNASSIGNED_ROW_COL_INDEX) and Assigned(Worksheet) then
    begin
      gcLastUsed := GetGridCol(scLastUsed);
      while (gcLast < ColCount-1) and (gcLast < gcLastUsed) do begin
        inc(gcLast);
        cell := Worksheet.FindCell(sr, GetWorksheetCol(gcLast));
        // Empty cell --> proceed with next cell to the right
        if (cell = nil) or (cell^.ContentType = cctEmpty) or
           ((cell^.ContentType = cctUTF8String) and (cell^.UTF8StringValue = ''))
        then
          continue;
        // Overflow possible from non-merged, horizontal, non-left-aligned label cells
        fmt := Worksheet.GetPointerToEffectiveCellFormat(cell);
        if (not Worksheet.IsMerged(cell)) and
           (cell^.ContentType = cctUTF8String) and
           not (uffTextRotation in fmt^.UsedFormattingFields) and
           (uffHorAlign in fmt^.UsedFormattingFields) and (fmt^.HorAlignment <> haLeft)
        then
          Break;
        // All other cases --> no overflow --> return to initial right column
        gcLast := ALastCol;
        Break;
      end;
    end;

    // Here begins the drawing loop of all cells in the row between gc and gclast
    while (gc <= gcLast) do begin
      gr := ARow;
      rct := AClipRect;
      // FDrawingCell is the cell which is currently being painted. We store
      // it to avoid excessive calls to "FindCell".
      FDrawingCell := nil;
      gcNext := gc + 1;
      if Assigned(Worksheet) and (gr >= FHeaderCount) and (gc >= FHeaderCount) then
      begin
        cell := Worksheet.FindCell(GetWorksheetRow(gr), GetWorksheetCol(gc));
        if (cell = nil) or (not Worksheet.IsMerged(cell)) then
        begin
          // single cell
          FDrawingCell := cell;
          if Worksheet.HasComment(cell) then
            commentcell_rct := CellRect(gc, gr)
          else
            commentcell_rct := Rect(0,0,0,0);
          // Special treatment of overflowing cells
          if FTextOverflow then
          begin
            gds := GetGridDrawState(gc, gr);
            ColRowToOffset(true, true, gc, rct.Left, rct.Right);
            if CellOverflow(gc, gr, gds, gc1, gc2, rct) then
            begin
              // Draw individual cells of the overflown range
              if IsRightToLeft then begin
                ColRowToOffset(true, true, gc1, tmp, rct.Right);
                ColRowToOffset(true, true, gc2, rct.Left, tmp);
              end else begin
                ColRowToOffset(true, true, gc1, rct.Left, tmp);    // rct is the clip rect
                ColRowToOffset(true, true, gc2, tmp, rct.Right);
              end;
              FDrawingCell := nil;
              temp_rct := rct;
              for i:= gc2 downto gc1 do begin
              // Starting from last col will ensure drawing grid lines below text
              // when text is overflowing in RTL, no problem in LTR
              // (Modification by "shobits1" - ok)
                ColRowToOffset(true, true, i, temp_rct.Left, temp_rct.Right);
                if HorizontalIntersect(temp_rct, clipArea) and (i <> gc) then
                begin
                  gds := GetGridDrawState(i, gr);
                  InternalDrawCell(i, gr, rct, temp_rct, gds);
                end;
              end;
              // Repaint the base cell text (it was partly overwritten before)
              FDrawingCell := cell;
              FTextOverflowing := true;
              ColRowToOffset(true, true, gc, temp_rct.Left, temp_rct.Right);
              if HorizontalIntersect(temp_rct, clipArea) then
              begin
                gds := GetGridDrawState(gc, gr);
                IntersectRect(rct, rct, clipArea);
                InternalDrawCell(gc, gr, rct, temp_rct, gds);
                if Worksheet.HasComment(FDrawingCell) then
                  DrawCommentMarker(temp_rct);
              end;
              FTextOverflowing := false;

              gcNext := gc2 + 1;
              gc := gcNext;
              continue;
            end;
          end;
        end
        else
        begin
          // merged cells
          FDrawingCell := Worksheet.FindMergeBase(cell);
          if Worksheet.FindMergedRange(FDrawingCell, sr1, sc1, sr2, sc2) then
          begin
            gr := GetGridRow(sr1);
            if Worksheet.HasComment(FDrawingCell) then
              commentcell_rct := CellRect(GetGridCol(sc2), gr)
            else
              commentcell_rct := Rect(0,0,0,0);
            ColRowToOffSet(False, True, gr, rct.Top, tmp);
            ColRowToOffSet(False, True, gr + integer(sr2) - integer(sr1), tmp, rct.Bottom);
            gc := GetGridCol(sc1);
            gcNext := gc + (sc2 - sc1) + 1;
          end;
        end;
      end;

      // Take care of upper and lower bounds of merged cells!
      temp_rct := rct;
      rct := CellRect(gc, gr, gcNext-1, gr);
      rct.Top := temp_rct.Top;
      rct.Bottom := temp_rct.Bottom;

      if (rct.Left < rct.Right) and HorizontalIntersect(rct, clipArea) then
      begin
        // Define clipping rectangle in order to avoid painting into the header cells
        clip_rct := rct;
        clip_rct.Top := AClipRect.Top;
        clip_rct.Bottom := AClipRect.Bottom;
        if (gc >= FHeaderCount + FFrozenCols) then begin
          if UseRightToLeftReading then begin
            if (clip_rct.Right > FTopLeft.X) then
              clip_rct.Right := FTopLeft.X;
          end else
            if (clip_rct.Left < FTopLeft.X) then
              clip_rct.Left := FTopLeft.X;
        end;
        gds := GetGridDrawState(gc, gr);

        // Draw cell
        InternalDrawCell(gc, gr, clip_rct, rct, gds);
        // Draw comment marker
        if (commentcell_rct.Left <> 0) and (commentcell_rct.Right <> 0) and
           (commentcell_rct.Top <> 0) and (commentcell_rct.Bottom <> 0)
        then
          DrawCommentMarker(commentcell_rct);
      end;

      gc := gcNext;
    end;
  end;    // with GCache.VisibleGrid ...

end;

{@@ ----------------------------------------------------------------------------
  Internal general text drawing method.

  @param  AText           Text to be drawn
  @param  AMeasureText    Text used for checking if the text fits into the
                          text rectangle. If too large and ReplaceTooLong = true,
                          a series of # is drawn.
  @param  ARect           Rectangle in which the text is drawn
  @param  AJustification  Determines whether the text is drawn at the "start" (0),
                          "center" (1) or "end" (2) of the drawing rectangle.
                          Start/center/end are seen along the text drawing
                          direction.
  @param ACellHorAlign    Is the HorAlignment property stored in the cell
  @param ACellVertAlign   Is the VertAlignment property stored in the cell
  @param ATextRot         Determines the rotation angle of the text.
  @param ATextWrap        Determines if the text can wrap into multiple lines
  @param AFontIndex       Font index to be used for drawing non-rich-text.
  @param ARichTextParams  an array of character and font index combinations for
                          rich-text formatting of text in cell
  @param AIsRightToLeft   if true cell must be drawn in right-to-left mode.

  @Note The reason to separate AJustification from ACellHorAlign and ACelVertAlign is
  the output of nfAccounting formatted numbers where the numbers are always
  right-aligned, and the currency symbol is left-aligned.
  THIS FEATURE IS CURRENTLY NO LONGER SUPPORTED.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.InternalDrawTextInCell(AText: String;
  ARect: TRect; ACellHorAlign: TsHorAlignment; ACellVertAlign: TsVertAlignment;
  ATextRot: TsTextRotation; ATextWrap: Boolean; AFontIndex: Integer;
  AOverrideTextColor: TColor; ARichTextParams: TsRichTextParams;
  AIsRightToLeft: Boolean);
begin
  // Since - due to the rich-text mode - characters are drawn individually their
  // background occasionally overpaints the prev characters (italic). To avoid
  // this we do not paint the character background - it is not needed anyway.
  Canvas.Brush.Style := bsClear;

  // Work horse for text drawing, both standard text and rich-text
  DrawRichText(Canvas, Workbook, ARect, AText, ARichTextParams, AFontIndex,
    ATextWrap, ACellHorAlign, ACellVertAlign, ATextRot, AOverrideTextColor,
    AIsRightToLeft, ZoomFactor
  );
end;

{@@ ----------------------------------------------------------------------------
  Standard key handling method inherited from TCustomGrid. Is overridden to
  catch some keys for special processing.

  @param  Key    Key which has been pressed
  @param  Shift  Additional shift keys which are pressed
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.KeyDown(var Key : Word; Shift : TShiftState);
var
  R: TRect;
begin
  case Key of
    VK_RIGHT:
      if (aeNavigation in FAutoExpand) and (Col = ColCount-1) then
        ColCount := ColCount + 1;
    VK_DOWN:
      if (aeNavigation in FAutoExpand) and (Row = RowCount-1) then
        RowCount := RowCount + 1;
    VK_END:
      if (aeNavigation in FAutoExpand) and (Col = ColCount-1) then
      begin
        R := GCache.FullVisibleGrid;
        ColCount := ColCount + R.Right - R.Left;
      end;
    VK_NEXT:  // Page down
      if (aeNavigation in FAutoExpand) and (Row = RowCount-1) then
      begin
        R := GCache.FullVisibleGrid;
        RowCount := Row + R.Bottom - R.Top;
      end;
    VK_F2:
      FEnhEditMode := true;
  end;

  inherited;

  case Key of
    VK_C, VK_X, VK_V:
      if Shift = [ssCtrl] then Key := 0;
      // Clipboard has already been handled, avoid passing key to CellAction
  end;
end;

                               (*
{@@ ----------------------------------------------------------------------------
  Loads the worksheet into the grid and displays its contents.

  @param   AWorksheet   Worksheet to be displayed in the grid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadFromWorksheet(AWorksheet: TsWorksheet);
begin
  if FWorkbookSource <> nil then
    exit;

  GetWorkbookSource.LoadFro
  FOwnedWorksheet := AWorksheet;
  if FOwnedWorksheet <> nil then begin
    inc(FLockSetup);
    FOwnedWorksheet.OnChangeCell := @ChangedCellHandler;
    FOwnedWorksheet.OnChangeFont := @ChangedFontHandler;
    ShowHeaders := (soShowHeaders in Worksheet.Options);
    ShowGridLines := (soShowGridLines in Worksheet.Options);
    if (soHasFrozenPanes in Worksheet.Options) then begin
      FrozenCols := FOwnedWorksheet.LeftPaneWidth;
      FrozenRows := FOwnedWorksheet.TopPaneHeight;
    end else begin
      FrozenCols := 0;
      FrozenRows := 0;
    end;
    Row := FrozenRows;
    Col := FrozenCols;
    dec(FLockSetup);
  end;
  Setup;
end;
                                 *)
{@@ ----------------------------------------------------------------------------
  Creates a new workbook and loads the given file into it. The file is assumed
  to have the given file format. Shows the sheet with the given sheet index.

  Call this method only for built-in file formats.

  @param   AFileName        Name of the file to be loaded
  @param   AFormat          Spreadsheet file format assumed for the file
  @param   AWorksheetIndex  Index of the worksheet to be displayed in the grid
                            (If empty then the active worksheet is loaded)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AFormat: TsSpreadsheetFormat; AWorksheetIndex: Integer);
var
  ae: TsAutoExpandModes;
begin
  ae := RelaxAutoExpand;
  GetWorkbookSource.LoadFromSpreadsheetFile(AFileName, AFormat, AWorksheetIndex);
  RestoreAutoExpand(ae);
end;

{@@ ----------------------------------------------------------------------------
  Creates a new workbook and loads the given file into it. The file is assumed
  to have the given file format. Shows the sheet with the given sheet index.

  Call this method for both built-in and user-provided file formats.

  @param   AFileName        Name of the file to be loaded
  @param   AFormatID        Spreadsheet file format identifier assumed for the
                            file (automatic detection if empty)
  @param   AWorksheetIndex  Index of the worksheet to be displayed in the grid
                            (If empty then the active worksheet is loaded)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadFromSpreadsheetFile(AFileName: string;
  AFormatID: TsSpreadFormatID = sfidUnknown; AWorksheetIndex: Integer = -1);
var
  ae: TsAutoExpandModes;
begin
  ae := RelaxAutoExpand;
  GetWorkbookSource.LoadFromSpreadsheetFile(AFileName, AFormatID, AWorksheetIndex);
  RestoreAutoExpand(ae);
end;

{@@ ----------------------------------------------------------------------------
  Creates a new workbook and loads the given file into it. Shows the sheet
  with the specified sheet index. The file format is determined automatically.

  @param   AFileName        Name of the file to be loaded
  @param   AWorksheetIndex  Index of the worksheet to be shown in the grid
                            (If empty then the active worksheet is loaded)
  @param   AFormatID        Spreadsheet file format identifier assumed for the
                            file (automatic detection if empty)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadSheetFromSpreadsheetFile(AFileName: String;
  AWorksheetIndex: Integer = -1; AFormatID: TsSpreadFormatID = sfidUnknown);
var
  ae: TsAutoExpandModes;
begin
  ae := RelaxAutoExpand;
  GetWorkbookSource.LoadFromSpreadsheetFile(AFilename, AFormatID, AWorksheetIndex);
  RestoreAutoExpand(ae);
end;

{@@ ----------------------------------------------------------------------------
  Loads an existing workbook into the grid.

  @param   AWorkbook        Workbook that had been created/loaded before.
  @param   AWorksheetIndex  Index of the worksheet to be shown in the grid
                            (If empty then the active worksheet is loaded)

  @Note THE CALLING PROCEDURE MUST NOT DESTROY THE WORKBOOK! The workbook will
   be destroyed by the workbook source.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.LoadFromWorkbook(AWorkbook: TsWorkbook;
  AWorksheetIndex: Integer = -1);
var
  ae: TsAutoExpandModes;
begin
  ae := RelaxAutoExpand;
  GetWorkbookSource.LoadFromWorkbook(AWorkbook, AWorksheetIndex);
  RestoreAutoExpand(ae);
  Invalidate;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookLink telling which item of the
  spreadsheet has changed.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
var
  actgrow, actgcol: Integer;
  grow, gcol: Integer;
  srow, scol: Cardinal;
  cell: PCell;
  lRow: PRow;

  {$IFDEF GRID_DEBUG}
  procedure DebugNotification(ACaption: String);
  var
    s: String;
  begin
    WriteLn(ACaption);
    s := '';
    if (lniWorksheet in AChangedItems) then s := s + 'lniWorksheet, ';
    if (lniCell in AChangedItems) then s := s + 'lniCell, ';
    if (lniSelection in AChangedItems) then s := s + 'lniSelection, ';
    if (lniAbortSelection in AChangedItems) then s := s + 'lniAbortSelection, ';
    if (lniRow in AChangedItems) then s := s + 'lniRow, ';
    if (lniCol in AChangedItems) then s := s + 'lniCol, ';
    if (lniWorksheetZoom in AChangedItems) then s := s + 'lniWorksheetZoom, ';
    if s <> '' then SetLength(s, Length(s) - 2);
    WriteLn('  AChangedItems = [', s, ']');
    WriteLn('  ActiveCellRow: ', Worksheet.ActiveCellRow, ' ActiveCellCol: ', Worksheet.ActiveCellCol);
    WriteLn('  TopRow: ', Worksheet.TopRow, ' LeftCol: ', Worksheet.LeftCol);
    WriteLn;
  end;
  {$ENDIF}

begin
  Unused(AData);

  {$IFDEF GRID_DEBUG}
  if Worksheet <> nil then
    DebugNotification('BEFORE ListenerNotification WorksheetGrid "' + Worksheet.Name + '":');
  {$ENDIF}

  // Nothing to do for  "workbook changed" because this is always combined with
  // "worksheet changed".

  // Worksheet changed
  if (lniWorksheet in AChangedItems) then
  begin
    BeginUpdate;   // avoid flicker...
    try
      if (Worksheet <> nil) then
      begin
        // remember indexes of top/left and active cell
        grow := GetGridRow(Worksheet.TopRow);
        gcol := GetGridCol(Worksheet.LeftCol);
        actgrow := GetGridRow(Worksheet.ActiveCellRow);
        actgcol := GetGridCol(Worksheet.ActiveCellCol);
        AutoExpandToRow(grow, aeNavigation);
        AutoExpandToCol(gcol, aeNavigation);
        if (grow <> Row) or (gcol <> Col) then
          MoveExtend(false, gcol, grow);
        inc(FLockSetup);
        // Setup grid headers and col/row count
        ShowHeaders := (soShowHeaders in Worksheet.Options);
        ShowGridLines := (soShowGridLines in Worksheet.Options);
        if (soHasFrozenPanes in Worksheet.Options) then begin
          FrozenCols := Worksheet.LeftPaneWidth;
          FrozenRows := Worksheet.TopPaneHeight;
        end else begin
          FrozenCols := 0;
          FrozenRows := 0;
        end;
        case Worksheet.BiDiMode of
          bdDefault: ParentBiDiMode := true;
          bdLTR    : begin
                       ParentBiDiMode := false;
                       BiDiMode := bdLeftToRight;
                     end;
          bdRTL    : begin
                       ParentBiDiMode := false;
                       BiDiMode := bdRightToLeft;
                     end;
        end;
        dec(FLockSetup);
      end;
      Setup;
      // scroll the grid for top/left to be as stored in the sheet
      if (grow <> TopRow) or (gcol <> LeftCol) then
      begin
        TopRow := gRow;
        LeftCol := gCol;
      end;
      // Select active cell
      AutoExpandToRow(actgrow, aeNavigation);
      AutoExpandToCol(actgcol, aeNavigation);
      if (actgrow <> Row) or (actgcol <> Col) then
        MoveExtend(false, actgcol, actgrow);
    finally
      EndUpdate;
    end;
  end;

  // Cell value or format changed
  if (lniCell in AChangedItems) then
  begin
    cell := PCell(AData);
    if (cell <> nil) then begin
      grow := GetGridRow(cell^.Row);
      gcol := GetGridCol(cell^.Col);
      AutoExpandToRow(grow, aeData);
      AutoExpandToCol(gcol, aeData);
      lRow := Worksheet.FindRow(cell^.Row);
      if (lRow = nil) or (lRow^.RowHeightType <> rhtCustom) then
        UpdateRowHeight(grow, true);
    end;
    Invalidate;
  end;

  // Selection changed
  if (lniSelection in AChangedItems) and (Worksheet <> nil) then
  begin
    grow := GetGridRow(Worksheet.ActiveCellRow);
    gcol := GetGridCol(Worksheet.ActiveCellCol);
    AutoExpandToRow(grow, aeNavigation);
    AutoExpandToCol(gcol, aeNavigation);
    if (grow <> Row) or (gcol <> Col) then
      MoveExtend(false, gcol, grow);
    if Worksheet.IsProtected then
    begin
      cell := Worksheet.FindCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
      FReadOnly := (cell = nil) or (cpLockCell in Worksheet.ReadCellProtection(cell));
    end else
      FReadOnly := false;
  end;

  // Abort selection because of an error
  if (lniAbortSelection in AChangedItems) and (Worksheet <> nil) then
  begin
    MouseUp(mbLeft, [], GCache.ClickMouse.X, GCache.ClickMouse.Y);
    // HOW TO DO THIS????    SelectActive not working...
  end;

  // Row height (after font or row record change).
  if (lniRow in AChangedItems) and (Worksheet <> nil) then
  begin
    srow := {%H-}PtrInt(AData);  // sheet row
    grow := GetGridRow(srow);    // grid row
    AutoExpandToRow(grow, aeData);
    lRow := Worksheet.FindRow(srow);
    if (lRow = nil) or (lRow^.RowHeightType <> rhtCustom) then
      UpdateRowHeight(grow, true);
  end;

  // Column width
  if (lniCol in AChangedItems) and (Worksheet <> nil) then
  begin
    scol := {%H-}PtrInt(AData);  // sheet column index
    gcol := GetGridCol(scol);
    //lCol := Worksheet.FindCol(scol);
    UpdateColWidth(gcol);
  end;

  // Worksheet zoom
  if (lniWorksheetZoom in AChangedItems) and (Worksheet <> nil) then
    AdaptToZoomFactor; // Reads value directly from Worksheet

  {$IFDEF GRID_DEBUG}
  if Worksheet <> nil then
    DebugNotification('AFTER ListenerNotification WorksheetGrid "' + Worksheet.Name + '":');
  {$ENDIF}

end;

{@@ ----------------------------------------------------------------------------
  Merges the selected cells to a single large cell
  Only the upper left cell can have content and formatting (which is extended
  into the other cells).
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MergeCells;
begin
  MergeCells(Selection);
end;

{@@ ----------------------------------------------------------------------------
  Merges the cells of the specified cell block to a single large cell
  Only the upper left cell can have content and formatting (which is extended
  into the other cells).
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MergeCells(ARect: TGridRect);
begin
  MergeCells(ARect.Left, ARect.Top, ARect.Right, ARect.Bottom);
end;

{@@ ----------------------------------------------------------------------------
  Merges the cells of the specified cell block to a single large cell
  Only the upper left cell can have content and formatting (which is extended
  into the other cells).
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MergeCells(ALeft, ATop, ARight, ABottom: Integer);
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Worksheet.MergeCells(
    GetWorksheetRow(ATop),
    GetWorksheetCol(ALeft),
    GetWorksheetRow(ABottom),
    GetWorksheetCol(ARight)
  );
end;

{@@ ----------------------------------------------------------------------------
  Standard mouse down handler. Is overridden here to handle hyperlinks and to
  enter "enhanced edit mode" which removes formatting from the values and
  presents formulas for editing.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MouseDown(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
{todo: extend such that the hyperlink is handled only when the text is clicked
       (tough because of overflow cells!) }
var
  mouseCell: TPoint;
  cell: PCell;
  r, c: Cardinal;
begin
  if FAllowDragAndDrop and
     (not Assigned(DragManager) or not DragManager.IsDragging) and
     (ssLeft in Shift) and
     MouseOnCellBorder(Point(X, Y), Selection) then
  begin
    if DragBorderBitmap = nil then
      CreateFillPattern(DragBorderBitmap, fsGray50, clBlack, clWhite);
    FDragStartCol := Col;
    FDragStartRow := Row;
    FOldDragStartCol := Col;
    FOldDragStartRow := Row;
    BeginDrag(false, DragManager.DragThreshold);
    exit;
  end;

  inherited;

  if (ssLeft in Shift) then
  begin
    { Prepare processing of the hyperlink: triggers a timer, the hyperlink is
      executed when the timer has expired (see HyperlinkTimerElapsed). }
    mouseCell := MouseToCell(Point(X, Y));
    r := GetWorksheetRow(mouseCell.Y);
    c := GetWorksheetCol(mouseCell.X);
    cell := Worksheet.FindCell(r, c);
    if Worksheet.IsMerged(cell) then
      cell := Worksheet.FindMergeBase(cell);
    if Worksheet.HasHyperlink(cell) then
    begin
      FHyperlinkCell := cell;
      FHyperlinkTimer.Enabled := true;
    end else
    begin
      FHyperlinkCell := nil;
      FHyperlinkTimer.Enabled := false;
    end;
  end;

  FEnhEditMode := true;
end;

{@@ ----------------------------------------------------------------------------
  Standard mouse move handler. Is overridden because, if TextOverflow is active,
  overflown cell may be erased when the mouse leaves them; repaints entire
  grid instead.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  prevMouseCell: TPoint;
begin
  prevMouseCell := GCache.MouseCell;

  inherited;

  if FTextOverflow and
     ((prevMouseCell.X <> GCache.MouseCell.X) or (prevMouseCell.Y <> GCache.MouseCell.Y))
  then
    InvalidateGrid;

  if Assigned(Dragmanager) and DragManager.IsDragging then
  begin;
    Cursor := crDefault;
  end else
  begin
    if FHyperlinkTimer.Enabled and (ssLeft in Shift) then
      FHyperlinkTimer.Enabled := false;

    if MouseOnCellBorder(Point(X, Y), Selection) then
      Cursor := crSize
    else
      Cursor := crDefault;
  end;
end;

{@@ Checks whether the specified point is on the border of a given cell.
  The tolerance is defined by the global variable CELL_BORDER_DELTA }
function TsCustomWorksheetGrid.MouseOnCellBorder(const APoint: TPoint;
  const ACellRect: TGridRect): Boolean;
var
  R: TRect;
  R1, R2: TRect;
begin
  R := CellRect(ACellRect.Left, ACellRect.Top, ACellRect.Right, ACellRect.Bottom);
  R1 := R;
  InflateRect(R1, CELL_BORDER_DELTA, CELL_BORDER_DELTA);
  R2 := R;
  InflateRect(R2, -CELL_BORDER_DELTA, -CELL_BORDER_DELTA);
  Result := PtInRect(R1, APoint) and not PtInRect(R2, APoint);
end;

procedure TsCustomWorksheetGrid.MouseUp(Button: TMouseButton;
  Shift:TShiftState; X,Y:Integer);
begin
  if FHyperlinkTimer.Enabled then begin
    FHyperlinkTimer.Enabled := false;
    FHyperlinkCell := nil;
  end;

  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Standard method inherited from TCustomGrid.
  Notifies the WorkbookSource of the changed selected cell.
  Repaints the grid after moving selection to avoid spurious rests of the
  old thick selection border.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.MoveSelection;
var
  sel: TsCellRangeArray;
  {$IFNDEF FPS_NO_GRID_MULTISELECT}
  i: Integer;
  {$ENDIF}
begin
  if (FActiveCellLock > 0) or (Assigned(DragManager) and DragManager.IsDragging) then
    exit;

  if Worksheet <> nil then
  begin
    {$IFNDEF FPS_NO_GRID_MULTISELECT}
    if HasMultiSelection then
    begin
      SetLength(sel, SelectedRangeCount);
      for i:=0 to High(sel) do
        with SelectedRange[i] do
        begin
          sel[i].Row1 := GetWorksheetRow(Top);
          sel[i].Col1 := GetWorksheetCol(Left);
          sel[i].Row2 := GetWorksheetRow(Bottom);
          sel[i].Col2 := GetWorksheetCol(Right);
        end;
    end else
    begin
      SetLength(sel, 1);
      sel[0].Row1 := GetWorksheetRow(Selection.Top);
      sel[0].Col1 := GetWorksheetCol(Selection.Left);
      sel[0].Row2 := GetWorksheetRow(Selection.Bottom);
      sel[0].Col2 := GetWorksheetRow(Selection.Right);
    end;
    {$ELSE}
    SetLength(sel, 1);
    sel[0].Row1 := GetWorksheetRow(Selection.Top);
    sel[0].Col1 := GetWorksheetCol(Selection.Left);
    sel[0].Row2 := GetWorksheetRow(Selection.Bottom);
    sel[0].Col2 := GetWorksheetRow(Selection.Right);
    {$ENDIF}
    Worksheet.SetSelection(sel);

    Worksheet.SelectCell(GetWorksheetRow(Row), GetWorksheetCol(Col));
  end;
  inherited;
  Refresh;
end;

{@@ ----------------------------------------------------------------------------
  Creates a new empty workbook with the specified number of columns and rows.

  @param   AColCount   Number of columns
  @param   ARowCount   Number of rows
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.NewWorkbook(AColCount, ARowCount: Integer);
begin
  GetWorkbookSource.CreateNewWorkbook;
  ColCount := AColCount + FHeaderCount;
  RowCount := ARowCount + FHeaderCount;
  Setup;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification: The grid is notified that the WorkbookLink
  is being removed.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ Prepares the Canvas default font for methods determining text size }
procedure TsCustomWorksheetGrid.PrepareCanvasFont;
var
  fnt: TsFont;
begin
  if Worksheet = nil then
    Canvas.Font.Assign(Font)
  else
  begin
    fnt := Workbook.GetDefaultFont;
    Convert_sFont_to_Font(fnt, Canvas.Font);
  end;
  Canvas.Font.Height := Round(ZoomFactor * Canvas.Font.Height);
end;

function TsCustomWorksheetGrid.RelaxAutoExpand: TsAutoExpandModes;
begin
  Result := FAutoExpand;
  FAutoExpand := [aeData, aeNavigation];
end;

procedure TsCustomWorksheetGrid.RestoreAutoExpand(AValue: TsAutoExpandModes);
begin
  FAutoExpand := AValue;
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the WorksheetGrid to the WorkbookSource.
  Required before destruction.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Writes the workbook represented by the grid to a spreadsheet file.

  Call this method only for built-in file formats.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved.
  @param   AFormat            Spreadsheet file format in which the file is to be
                              saved.
  @param   AOverwriteExisting If the file already exists, it is overwritten in
                              the case of AOverwriteExisting = true, or an
                              exception is raised if AOverwriteExisting = false
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SaveToSpreadsheetFile(AFileName: String;
  AFormat: TsSpreadsheetFormat; AOverwriteExisting: Boolean = true);
begin
  if Workbook <> nil then
    Workbook.WriteToFile(AFileName, AFormat, AOverwriteExisting);
end;

{@@ ----------------------------------------------------------------------------
  Writes the workbook represented by the grid to a spreadsheet file.

  Call this method for both built-in and user-provided file formats.

  @param   AFileName          Name of the file to which the workbook is to be
                              saved.
  @param   AFormatID          Identifier for the spreadsheet file format in
                              which the file is to be saved.
  @param   AOverwriteExisting If the file already exists, it is overwritten in
                              the case of AOverwriteExisting = true, or an
                              exception is raised if AOverwriteExisting = false
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SaveToSpreadsheetFile(AFileName: String;
  AFormatID: TsSpreadFormatID; AOverwriteExisting: Boolean = true);
begin
  if Workbook <> nil then
    Workbook.WriteToFile(AFileName, AFormatID, AOverwriteExisting);
end;

{@@ ----------------------------------------------------------------------------
  Saves the workbook into a file with the specified file name. If this file
  name already exists the file is overwritten if AOverwriteExisting is true.

  @param   AFileName           Name of the file to which the workbook is to be
                               saved
                               If the file format is not known it is written
                               as BIFF8/XLS.
  @param   AOverwriteExisting  If this file already exists it is overwritten if
                               AOverwriteExisting = true, or an exception is
                               raised if AOverwriteExisting = false.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SaveToSpreadsheetFile(AFileName: String;
  AOverwriteExisting: Boolean = true);
begin
  if Workbook <> nil then
    Workbook.WriteToFile(AFileName, AOverwriteExisting);
end;

{@@ ----------------------------------------------------------------------------
  Loads the workbook into the grid and selects the sheet with the given index.
  "Selected" means here that the sheet is loaded into the grid.

  @param   AIndex   Index of the worksheet to be shown in the grid
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SelectSheetByIndex(AIndex: Integer);
begin
  GetWorkbookSource.SelectWorksheet(Workbook.GetWorksheetByIndex(AIndex));
end;

{@@ ----------------------------------------------------------------------------
  Standard method inherited from TCustomGrid. Is overridden to prevent
  selection of cells in a protected worksheet. Details depend on whether
  the elements spSelectLockedCells and/or spSelectUnlockedCells are included in
  the worksheet's set of protections, and whether the cell to be selected is
  locked or not.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.SelectCell(ACol, ARow: Integer): Boolean;
var
  cell: PCell;
  cp: TsCellprotections;
begin
  Result := inherited;
  if Result and Worksheet.IsProtected then
  begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    cp := Worksheet.ReadCellProtection(cell);
    if (cpLockCell in cp) and (spSelectLockedCells in Worksheet.Protection) then
      Result := false;
    if not (cpLockCell in cp) and (spSelectUnlockedCells in Worksheet.Protection) then
      Result := false;
  end;
end;

{@@ Event handler which fires when an element of the SelectionPen changes. }
procedure TsCustomWorksheetGrid.GenericPenChangeHandler(Sender: TObject);
begin
  InvalidateGrid;
end;

{@@ ----------------------------------------------------------------------------
  Standard method inherited from TCustomGrid. Fetches the text that is
  currently in the editor. It is not yet transferred to the worksheet because
  input will be checked only at the end of editing.

  @param  ACol    Grid column index of the cell being edited
  @param  ARow    Grid row index of the cell being edited
  @param  AValue  String which is currently in the cell editor
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SetEditText(ACol, ARow: Longint; const AValue: string);
begin
  FEditText := AValue;
  FEditing := true;
  inherited SetEditText(aCol, aRow, aValue);
end;

{@@ ----------------------------------------------------------------------------
  Helper method for setting up the rows and columns after a new workbook is
  loaded or created. Sets up the grid's column and row count, as well as the
  initial column widths and row heights.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Setup;
var
  maxColCount, maxRowCount: Integer;
begin
  if csLoading in ComponentState then
    exit;

  if FLockSetup > 0 then
    exit;

  if not HandleAllocated then //or (not Parent.HandleAllocated) then
    //Avoid crash when accessing the canvas, e.g. in GetDefaultHeaderColWidth
    exit;

  if (Worksheet = nil) or (Worksheet.GetCellCount = 0) then begin
    FixedCols := FFrozenCols + FHeaderCount;
    FixedRows := FFrozenRows + FHeaderCount;
    if ShowHeaders then begin
      PrepareCanvasFont;  // Applies the zoom factor
      ColWidths[0] := GetDefaultHeaderColWidth;
      RowHeights[0] := GetDefaultRowHeight;
    end;
    FTopLeft := CalcTopLeft(false);
  end else
  if Worksheet <> nil then begin
    maxColCount := IfThen(aeDefault in FAutoExpand, DEFAULT_COL_COUNT, 1);
    maxRowCount := IfThen(aeDefault in FAutoExpand, DEFAULT_ROW_COUNT, 1);
    ColCount := Max(GetGridCol(Worksheet.GetLastColIndex) + 1, maxColCount);
    RowCount := Max(GetGridRow(Worksheet.GetLastRowIndex) + 1, maxRowCount);
    (*
    if aeDefault in FAutoExpand then begin
      ColCount := Max(GetGridCol(Worksheet.GetLastColIndex)+1, DEFAULT_COL_COUNT); // + FHeaderCount;
      RowCount := Max(GetGridRow(Worksheet.GetLastRowIndex)+1, DEFAULT_ROW_COUNT); // + FHeaderCount;
    end else begin

// wp: next lines replaced by the following ones because of
// http://forum.lazarus.freepascal.org/index.php/topic,36770.msg246192.html#msg246192
// NOTE: VERY SENSITIVE LOCATION !!!

//      ColCount := Max(GetGridCol(WorkSheet.GetLastColIndex), 1) + FHeaderCount;
//      RowCount := Max(GetGridCol(Worksheet.GetLastRowIndex), 1) + FHeaderCount;
      ColCount := Max(GetGridCol(WorkSheet.GetLastColIndex)+1, 1); // + FHeaderCount;
      RowCount := Max(GetGridCol(Worksheet.GetLastRowIndex)+1, 1); // + FHeaderCount;
    end;
    *)
    FixedCols := FFrozenCols + FHeaderCount;
    FixedRows := FFrozenRows + FHeaderCount;
    if ShowHeaders then begin
      PrepareCanvasFont;
      ColWidths[0] := GetDefaultHeaderColWidth;
      RowHeights[0] := GetDefaultRowHeight;
    end;
    FTopLeft := CalcTopLeft(false);
  end;
  UpdateColWidths;
  UpdateRowHeights;
  //Invalidate;   // wp: really needed? Might cause flicker
end;

{@@ ----------------------------------------------------------------------------
  Setter to define the link to the workbook.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;

  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FInternalWorkbookSource.RemoveListener(self);

  if (AValue = FInternalWorkbookSource) or (AValue = nil) then
  begin
    FWorkbookSource := nil;
    FInternalWorkbookSource.AddListener(self);
  end else
  begin
    FWorkbookSource := AValue;
    FWorkbookSource.AddListener(self);
  end;

  if not (csDestroying in ComponentState) and Assigned(Parent) then
    ListenerNotification([lniWorksheet, lniSelection]);
end;

{@@ ----------------------------------------------------------------------------
  Shows cell borders for the cells in the range between columns ALeft and ARight
  and rows ATop and ABottom.
  The border of the block's left outer edge is defined by ALeftOuterStyle,
  that of the block's top outer edge by ATopOuterStyle, etc.
  Set the color of a border style to scNotDefined or scTransparent in order to
  hide the corresponding border line, or use the constant NO_CELL_BORDER.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.ShowCellBorders(ALeft, ATop, ARight, ABottom: Integer;
  const ALeftOuterStyle, ATopOuterStyle, ARightOuterStyle, ABottomOuterStyle,
  AHorInnerStyle, AVertInnerStyle: TsCellBorderStyle);

  function BorderVisible(const AStyle: TsCellBorderStyle): Boolean;
  begin
    Result := (AStyle.Color <> scNotDefined) and (AStyle.Color <> scTransparent);
  end;

  procedure ProcessBorder(ARow, ACol: Cardinal; ABorder: TsCellBorder;
    const AStyle: TsCellBorderStyle);
  var
    cb: TsCellBorders = [];
    cell: PCell;
  begin
    cell := Worksheet.FindCell(ARow, ACol);
    if cell <> nil then
      cb := Worksheet.ReadCellBorders(cell);
    if BorderVisible(AStyle) then
    begin
      Include(cb, ABorder);
      cell := Worksheet.WriteBorders(ARow, ACol, cb);
      Worksheet.WriteBorderStyle(cell, ABorder, AStyle);
    end else
    if cb <> [] then
    begin
      Exclude(cb, ABorder);
      cell := Worksheet.WriteBorders(ARow, ACol, cb);
    end;
    FixNeighborCellBorders(cell);
  end;

var
  r, c, r1, c1, r2, c2: Cardinal;
begin
  if Worksheet = nil then
    exit;

  // Preparations
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  r1 := GetWorksheetRow(ATop);
  r2 := GetWorksheetRow(ABottom);
  c1 := GetWorksheetCol(ALeft);
  c2 := GetWorksheetCol(ARight);

  // Top outer border
  for c := c1 to c2 do
    ProcessBorder(r1, c, cbNorth, ATopOuterStyle);
  // Bottom outer border
  for c := c1 to c2 do
    ProcessBorder(r2, c, cbSouth, ABottomOuterStyle);
  // Left outer border
  for r := r1 to r2 do
    ProcessBorder(r, c1, cbWest, ALeftOuterStyle);
  // Right outer border
  for r := r1 to r2 do
    ProcessBorder(r, c2, cbEast, ARightOuterStyle);
  // Horizontal inner border
  if r1 <> r2 then
    for r := r1 to r2-1 do
      for c := c1 to c2 do
        ProcessBorder(r, c, cbSouth, AHorInnerStyle);
  // Vertical inner border
  if c1 <> c2 then
    for r := r1 to r2 do
      for c := c1 to c2-1 do
        ProcessBorder(r, c, cbEast, AVertInnerStyle);
end;

{@@ ----------------------------------------------------------------------------
  Sorts the grid by calling the corresponding method of the worksheet.
  Sorting extends across the entire worksheet.
  Sort direction is determined by the property "SortOrder". Other sorting
  criteria are "case-sensitive" and "numbers first".

  @param AColSorting If true the grid is sorted from top to bottom and the
                     next parameter, "Index", refers to a column. Otherweise
                     sorting goes from left to right and "Index" refers to a row.
  @param AIndex      Index of the column (if ColSorting=true) or row (ColSorting = false)
                     which is sorted.
  @param AIndxFrom   Sorting starts at this row (ColSorting=true) / column (ColSorting=false)
  @param AIndxTo     Sorting ends at this row (ColSorting=true) / column (ColSorting=false)
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.Sort(AColSorting: Boolean;
  AIndex, AIndxFrom, AIndxTo:Integer);
var
  sortParams: TsSortParams;
begin
  sortParams := InitSortParams(AColSorting, 1);
  sortParams.Keys[0].ColRowIndex := AIndex - HeaderCount;
  if SortOrder = soDescending then
    sortParams.Keys[0].Options := [ssoDescending];
  if AColSorting then
    Worksheet.Sort(
      sortParams,
      AIndxFrom-HeaderCount, 0, AIndxTo-HeaderCount, Worksheet.GetLastColIndex
    )
  else
    Worksheet.Sort(
      sortParams,
      0, AIndxFrom-HeaderCount, Worksheet.GetLastRowIndex, AIndxTo-HeaderCount
    );
end;

{@@ ----------------------------------------------------------------------------
  Converts a coordinate given in workbook units to pixels using the current
  screen resolution
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.ToPixels(AValue: Double): Integer;
var
  inches: Double;
begin
  inches := Workbook.ConvertUnits(AValue, Workbook.Units, suInches);
  Result := round(inches * Screen.PixelsPerInch);
end;

{@@ ----------------------------------------------------------------------------
  Store the value of the TopLeft cell in the worksheet
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.TopLeftChanged;
begin
  inherited;
  Worksheet.ScrollTo(GetWorkSheetRow(TopRow), GetWorksheetCol(LeftCol));
end;

{@@ ----------------------------------------------------------------------------
  Modifies the text that is show for cells which are too narrow to hold the
  entire text. The method follows the behavior of Excel and Open/LibreOffice:
  If the specified cell contains a non-formatted number, then it is formatted
  such that the text fits into the cell. If the text is still too long or
  the cell does not contain a label then the cell is filled by '#' characters.
  Label cell texts are not modified, they can overflow into the adjacent cells.
-------------------------------------------------------------------------------}
function TsCustomWorksheetGrid.TrimToCell(ACell: PCell): String;
var
  cellSize, txtSize: Integer;
  decs: Integer;
  p: Integer;
  isRotated: Boolean;
  isStacked: Boolean;
  fmt: PsCellFormat;
  numFmt: TsNumFormatParams;
  nfs: String;
  isGeneralFmt: Boolean;
  r1,c1,r2,c2: Cardinal;
begin
  Result := Worksheet.ReadAsText(ACell);
  if (Result = '') or ((ACell <> nil) and (ACell^.ContentType = cctUTF8String)) then
    exit;

//  fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
  fmt := Worksheet.GetPointerToEffectiveCellFormat(ACell^.Row, ACell^.Col);
  isRotated := (fmt^.TextRotation <> trHorizontal);
  isStacked := (fmt^.TextRotation = rtStacked);
  numFmt := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);
  isGeneralFmt := (numFmt = nil) or (numFmt.NumFormat = nfGeneral);

  // Determine space available in cell
  if Worksheet.IsMerged(ACell) then
  begin
    Worksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    cellSize := 0;
    if isRotated then
      for p:=GetGridRow(r1) to GetGridRow(r2) do cellSize := cellSize + RowHeights[p]
    else
      for p:=GetGridCol(c1) to GetGridCol(c2) do cellSize := cellSize + ColWidths[p];
  end else
  begin
    if isRotated then
      cellSize := RowHeights[GetGridRow(ACell^.Row)]
    else
      cellSize := ColWidths[GetGridCol(ACell^.Col)];
  end;
  cellSize := cellSize - 2*ConstCellPadding;

  // Determine space needed for text
  if isStacked then
    txtSize := Length(Result) * Canvas.TextHeight('A')
  else
    txtSize := Canvas.TextWidth(Result);

  // Nothing to do if text fits into cell
  if txtSize <= cellSize then
    exit;

  if (ACell^.ContentType = cctNumber) and isGeneralFmt then
  begin
    // Determine number of decimal places
    p := pos(Workbook.FormatSettings.DecimalSeparator, Result);
    if p = 0 then
      decs := 0
    else
      decs := Length(Result) - p;

    // Use floating point format, but reduce number of decimal places until
    // text fits in
    while decs > 0 do
    begin
      dec(decs);
      Result := Format('%.*f', [decs, ACell^.NumberValue], Workbook.FormatSettings);
      if isStacked then
        txtSize := Length(Result) * Canvas.TextHeight('A')
      else
        txtSize := Canvas.TextWidth(Result);
      if txtSize <= cellSize then
        exit;
    end;

    // There seem to be too many integer digits. Switch to exponential format.
    decs := 13;
    while decs > 0 do
    begin
      dec(decs);
      nfs := IfThen(decs = 0, '0E-00', '0.' + DupeString('0', decs) + 'E-00');
      Result := FormatFloat(nfs, ACell^.NumberValue, Workbook.FormatSettings);
//      Result := Format('%.*e', [decs, ACell^.NumberValue], Workbook.FormatSettings);
      if isStacked then
        txtSize := Length(Result) * Canvas.TextHeight('A')
      else
        txtSize := Canvas.TextWidth(Result);
      if txtSize <= cellSize then
        exit;
    end;
  end;

  // Still text too long or non-number --> Fill with # characters.
  Result := '';
  txtSize := 0;
  while txtSize < cellSize do
  begin
    Result := Result + '#';
    if isStacked then
      txtSize := Length(Result) * Canvas.TextHeight('#')
    else
      txtSize := Canvas.TextWidth(Result);
  end;

  // We added one character too many
  Delete(Result, Length(Result), 1);
end;

{@@ ----------------------------------------------------------------------------
  Splits a merged cell block into single cells
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UnmergeCells;
begin
  Worksheet.UnmergeCells(
    GetWorksheetRow(Selection.Top),
    GetWorksheetCol(Selection.Left)
  );
end;

{@@ ----------------------------------------------------------------------------
  If the specified cell belongs to a merged block, the merged block is
  split into single cells
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UnmergeCells(ACol, ARow: Integer);
begin
  Worksheet.UnmergeCells(
    GetWorksheetRow(ARow),
    GetworksheetCol(ACol)
  );
end;

procedure TsCustomWorksheetGrid.UpdateColWidth(ACol: Integer);
var
  lCol: PCol;
  w: Integer;       // Col width at current zoom level
  w100: Integer;    // Col width at 100% zoom level
begin
  if Worksheet <> nil then
  begin
    lCol := Worksheet.FindCol(ACol - FHeaderCount);
    if (lCol <> nil) and (lCol^.ColWidthType = cwtCustom) then
      w100 := CalcColWidthFromSheet(lCol^.Width)
    else
      w100 := CalcColWidthFromSheet(Worksheet.ReadDefaultColWidth(Workbook.Units));
    w := round(w100 * ZoomFactor);
  end else
    w := DefaultColWidth;   // Zoom factor has already been applied by getter
  ColWidths[ACol] := w;
end;

{@@ ----------------------------------------------------------------------------
  Updates column widths according to the data in the TCol records
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UpdateColWidths(AStartIndex: Integer = 0);
var
  i: Integer;
begin
  if AStartIndex = 0 then
    AStartIndex := FHeaderCount;
  BeginUpdate;
  try
    for i := AStartIndex to ColCount-1 do
      UpdateColWidth(i);
  finally
    EndUpdate;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Updates the height if the specified row in the grid with the value stored in
  the worksheet multiplied by the current zoom factor. If the stored row height
  type is rhtAuto (meaning: "row height is auto-calculated") and the current
  row height in the row record is 0 then the row height is calculated by
  iterating over all cells in this row. This happens also if the parameter
  AEnforceCalcRowHeight is true.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UpdateRowHeight(ARow: Integer;
  AEnforceCalcRowHeight: Boolean = false);
var
  lRow: PRow;
  sr: Cardinal;
  h: Integer;     // Row height, in pixels. Contains zoom factor.
  doCalcRowHeight: Boolean;
begin
  if ARow < FHeaderCount then
    exit;

  h := 0;
  if Worksheet <> nil then
  begin
    sr := ARow - FHeaderCount;    // worksheet row index
    lRow := Worksheet.FindRow(sr);
    if (lRow <> nil) then begin
      case lRow^.RowHeightType of
        rhtCustom:
          begin
            h := round(CalcRowHeightFromSheet(lRow^.Height) * ZoomFactor);
            if AEnforceCalcRowHeight then begin
              h := CalcAutoRowHeight(ARow);
              if h = 0 then begin
                h := DefaultRowHeight;
                lRow^.RowHeightType := rhtDefault;
              end else
                lRow^.RowHeightType := rhtAuto;
              lRow^.Height := CalcRowHeightToSheet(round(h / ZoomFactor));
            end;
          end;
        rhtAuto, rhtDefault:
          begin
            doCalcRowHeight := AEnforceCalcRowHeight or (lRow^.Height = 0);
            if doCalcRowHeight then begin
              // Calculate current grid row height in pixels by iterating over all cells in row
              h := CalcAutoRowHeight(ARow);  // ZoomFactor already applied to font heights
              if h = 0 then begin
                h := DefaultRowHeight;       // Zoom factor applied by getter function
                lRow^.RowHeightType := rhtDefault;
              end else
                lRow^.RowHeightType := rhtAuto;
              // Calculate the unzoomed row height in workbook units and store
              // in row record
              lRow^.Height := CalcRowHeightToSheet(round(h / ZoomFactor));
            end else
              // If autocalc mode is off we just take the row height from the row record
              h := round(CalcRowHeightFromSheet(lRow^.Height) * ZoomFactor);
          end;
      end;  // case
    end else
    // No row record so far.
    if Worksheet.GetCellCountInRow(sr) > 0 then
    begin
      // Case 1: This row does contain cells
      lRow := Worksheet.AddRow(sr);
      if AEnforceCalcRowHeight then
        h := CalcAutoRowHeight(ARow) else
        h := DefaultRowHeight;
      lRow^.Height := CalcRowHeightToSheet(round(h / ZoomFactor));
      if h <> DefaultRowHeight then
        lRow^.RowHeightType := rhtAuto
      else
        lRow^.RowHeightType := rhtDefault;
    end else
      // Case 2: No cells in row
      h := DefaultRowHeight;   // Zoom factor is applied by getter function
  end;

  if h = 0 then
    h := DefaultRowHeight;     // Zoom factor is applied by getter function

  inc(FZoomLock);  // We don't want to modify the sheet row heights here.
  RowHeights[ARow] := h;
  dec(FZoomLock);
end;

{@@ ----------------------------------------------------------------------------
  Updates grid row heights by using the data from the TRow records.
-------------------------------------------------------------------------------}
procedure TsCustomWorksheetGrid.UpdateRowHeights(AStartRow: Integer = -1;
  AEnforceCalcRowHeight: Boolean = false);
var
  r, r1: Integer;
begin
  if FRowHeightLock > 0 then
    exit;

  if AStartRow = -1 then
    r1 := FHeaderCount else
    r1 := AStartRow;

  BeginUpdate;
  try
    for r:=r1 to RowCount-1 do
      UpdateRowHeight(r, AEnforceCalcRowHeight);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.WMHScroll(var message: TLMHScroll);
begin
  inherited;
  //if Worksheet.GetImageCount > 0 then
    Invalidate;
end;

procedure TsCustomWorksheetGrid.WMVScroll(var message: TLMVScroll);
begin
  inherited;
  //if Worksheet.GetImageCount > 0 then
    Invalidate;
end;


{*******************************************************************************
*                      Setter / getter methods                                 *
*******************************************************************************}

function TsCustomWorksheetGrid.GetCellFontColor(ACol, ARow: Integer): TsColor;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := scNotDefined;
  if (Workbook <> nil) and (Worksheet <> nil) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    fnt := Worksheet.ReadCellFont(cell);
    Result := fnt.Color;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontColors(ALeft, ATop, ARight, ABottom: Integer): TsColor;
var
  c, r: Integer;
  clr: TsColor;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetCellFontColor(ALeft, ATop);
  clr := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetCellFontColor(c, r);
      if (Result <> clr) then begin
        Result := scNotDefined;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellFontName(ACol, ARow: Integer): String;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := '';
  if (Workbook <> nil) and (Worksheet <> nil) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    fnt := Worksheet.ReadCellFont(cell);
    if fnt <> nil then
      Result := fnt.FontName;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontNames(ALeft, ATop, ARight, ABottom: Integer): String;
var
  c, r: Integer;
  s: String;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetCellFontName(ALeft, ATop);
  s := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetCellFontName(c, r);
      if (Result <> '') and (Result <> s) then begin
        Result := '';
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellFontSize(ACol, ARow: Integer): Single;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := -1.0;
  if (Workbook <> nil) and (Worksheet <> nil) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    fnt := Worksheet.ReadCellFont(cell);
    Result := fnt.Size;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontSizes(ALeft, ATop, ARight, ABottom: Integer): Single;
var
  c, r: Integer;
  sz: Single;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetCellFontSize(ALeft, ATop);
  sz := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetCellFontSize(c, r);
      if (Result <> -1) and not SameValue(Result, sz, 1E-3) then begin
        Result := -1.0;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellFontStyle(ACol, ARow: Integer): TsFontStyles;
var
  cell: PCell;
  fnt: TsFont;
begin
  Result := [];
  if (Workbook <> nil) and (Worksheet <> nil) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    fnt := Worksheet.ReadCellFont(cell);
    Result := fnt.Style;
  end;
end;

function TsCustomWorksheetGrid.GetCellFontStyles(ALeft, ATop,
  ARight, ABottom: Integer): TsFontStyles;
var
  c, r: Integer;
  style: TsFontStyles;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetCellFontStyle(ALeft, ATop);
  style := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetCellFontStyle(c, r);
      if Result <> style then begin
        Result := [];
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetCellValue(ACol, ARow: Integer): variant;
var
  cell: PCell;
begin
  Result := Null;
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if cell <> nil then
      case cell^.ContentType of
        cctEmpty     : ;
        cctNumber,
        cctDateTime  : Result := cell^.NumberValue;
        cctUTF8String: Result := cell^.UTF8Stringvalue;
        cctBool      : Result := cell^.BoolValue;
        cctError     : Result := cell^.ErrorValue;
      end;
  end;
end;

function TsCustomWorksheetGrid.GetColWidths(ACol: Integer): Integer;
begin
  Result := inherited ColWidths[ACol];
end;

function TsCustomWorksheetGrid.GetDefColWidth: Integer;
begin
  Result := round(FDefColWidth100 * ZoomFactor);
end;

function TsCustomWorksheetGrid.GetDefRowHeight: Integer;
begin
  Result := round(FDefRowHeight100 * Zoomfactor);
end;

function TsCustomWorksheetGrid.GetHorAlignment(ACol, ARow: Integer): TsHorAlignment;
var
  cell: PCell;
begin
  Result := haDefault;
  if Assigned(Worksheet) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Result := Worksheet.ReadHorAlignment(cell);
  end;
end;

function TsCustomWorksheetGrid.GetHorAlignments(ALeft, ATop, ARight, ABottom: Integer): TsHorAlignment;
var
  c, r: Integer;
  horalign: TsHorAlignment;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetHorAlignment(ALeft, ATop);
  horalign := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetHorAlignment(c, r);
      if Result <> horalign then begin
        Result := haDefault;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetHyperlink(ACol, ARow: Integer): String;
var
  hlink: TsHyperlink;
begin
  Result := '';
  if Assigned(Worksheet) then
  begin
    hlink := Worksheet.ReadHyperLink(Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol)));
    if hlink.Target <> '' then begin
      Result := hlink.Target;
      if hlink.Tooltip <> '' then Result := Result + '|' + hlink.ToolTip;
    end;
  end;
end;

function TsCustomWorksheetGrid.GetNumberFormat(ACol, ARow: Integer): String;
var
  nf: TsNumberFormat;
  cell: PCell;
begin
  Result := '';
  if Assigned(Worksheet) and Assigned(Workbook) then
  begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if cell <> nil then
      Worksheet.ReadNumFormat(cell, nf, Result);
  end;
end;

function TsCustomWorksheetGrid.GetNumberFormats(ALeft, ATop,
  ARight, ABottom: Integer): String;
var
  c, r: Integer;
  nfs: String;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  nfs := GetNumberformat(ALeft, ATop);
  for r := ALeft to ARight do
    for c := ATop to ABottom do
      if nfs <> GetNumberFormat(c, r) then
      begin
        Result := '';
        exit;
      end;
  Result := nfs;
end;

function TsCustomWorksheetGrid.GetRowHeights(ARow: Integer): Integer;
begin
  Result := inherited RowHeights[ARow];
end;

function TsCustomWorksheetGrid.GetShowGridLines: Boolean;
begin
  Result := (Options * [goHorzLine, goVertLine] <> []);
end;

function TsCustomWorksheetGrid.GetShowHeaders: Boolean;
begin
  Result := FHeaderCount <> 0;
end;

function TsCustomWorksheetGrid.GetTextRotation(ACol, ARow: Integer): TsTextRotation;
var
  cell: PCell;
begin
  Result := trHorizontal;
  if Assigned(Worksheet) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Result := Worksheet.ReadTextRotation(cell);
  end;
end;

function TsCustomWorksheetGrid.GetTextRotations(ALeft, ATop,
  ARight, ABottom: Integer): TsTextRotation;
var
  c, r: Integer;
  textrot: TsTextRotation;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetTextRotation(ALeft, ATop);
  textrot := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetTextRotation(c, r);
      if Result <> textrot then begin
        Result := trHorizontal;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetWorkbookSource: TsWorkbookSource;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource else
    Result := FInternalWorkbookSource;
end;

function TsCustomWorksheetGrid.GetVertAlignment(ACol, ARow: Integer): TsVertAlignment;
var
  cell: PCell;
begin
  Result := vaDefault;
  if Assigned(Worksheet) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Result := Worksheet.ReadVertAlignment(cell);
  end;
end;

function TsCustomWorksheetGrid.GetVertAlignments(
  ALeft, ATop, ARight, ABottom: Integer): TsVertAlignment;
var
  c, r: Integer;
  vertalign: TsVertAlignment;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetVertalignment(ALeft, ATop);
  vertalign := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetVertAlignment(c, r);
      if Result <> vertalign then begin
        Result := vaDefault;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetWorkbook: TsWorkbook;
begin
  Result := GetWorkbookSource.Workbook;
end;

function TsCustomWorksheetGrid.GetWorksheet: TsWorksheet;
begin
  Result := GetWorkbookSource.Worksheet;
end;

function TsCustomWorksheetGrid.GetWordwrap(ACol, ARow: Integer): Boolean;
var
  cell: PCell;
begin
  Result := false;
  if Assigned(Worksheet) then begin
    cell := Worksheet.FindCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Result := Worksheet.ReadWordwrap(cell);
  end;
end;

function TsCustomWorksheetGrid.GetWordwraps(ALeft, ATop,
  ARight, ABottom: Integer): Boolean;
var
  c, r: Integer;
  wrapped: Boolean;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  Result := GetWordwrap(ALeft, ATop);
  wrapped := Result;
  for c := ALeft to ARight do
    for r := ATop to ABottom do begin
      Result := GetWordwrap(c, r);
      if Result <> wrapped then begin
        Result := false;
        exit;
      end;
    end;
end;

function TsCustomWorksheetGrid.GetZoomFactor: Double;
begin
  if Worksheet <> nil then
    Result := Worksheet.Zoomfactor
  else
    Result := 1.0;
end;

procedure TsCustomWorksheetGrid.SetAutoCalc(AValue: Boolean);
var
  optns: TsWorkbookOptions;
begin
  FAutoCalc := AValue;

  if Assigned(WorkbookSource) then
  begin
    optns := WorkbookSource.Options;
    if FAutoCalc then
      Include(optns, boAutoCalc) else
      Exclude(optns, boAutoCalc);
    WorkbookSource.Options := optns;
    if FInternalWorkbookSource <> nil then
      FInternalWorkbookSource.Options := optns;
  end;
end;

procedure TsCustomWorksheetGrid.SetBackgroundColor(ACol, ARow: Integer;
  AValue: TsColor);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then begin
    BeginUpdate;
    try
      cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
      Worksheet.WriteBackgroundColor(cell, AValue);
    finally
      EndUpdate;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetBackgroundColors(
  ALeft, ATop, ARight, ABottom: Integer; AValue: TsColor);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetBackgroundColor(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBiDiMode(ACol, ARow: Integer;
  AValue: TsBiDiMode);
begin
  if Assigned(Worksheet) then
    Worksheet.WriteBiDiMode(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetCellBorder(ACol, ARow: Integer;
  AValue: TsCellBorders);
var
  cell: PCell;
  sr1, sc1, sr2, sc2: Cardinal;
  gr1, gc1, gr2, gc2: Integer;
  styles, saved_styles: TsCellBorderStyles;
begin
  if Assigned(Worksheet) then begin
    BeginUpdate;
    try
      cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));

      if Worksheet.IsMergeBase(cell) then
      begin
        styles := Worksheet.ReadCellBorderStyles(cell);
        saved_styles := styles;
        if not (cbEast in AValue) then
          styles[cbEast] := NO_CELL_BORDER;
        if not (cbWest in AValue) then styles[cbWest] := NO_CELL_BORDER;
        if not (cbNorth in AValue) then styles[cbNorth] := NO_CELL_BORDER;
        if not (cbSouth in AValue) then styles[cbSouth] := NO_CELL_BORDER;
        Worksheet.FindMergedRange(cell, sr1, sc1, sr2, sc2);
        gr1 := GetGridRow(sr1);
        gr2 := GetGridRow(sr2);
        gc1 := GetGridCol(sc1);
        gc2 := GetGridCol(sc2);
        // Set border flags and styles for all outer cells of the merged block
        // Note: This overwrites the styles of the base ...
        ShowCellBorders(gc1,gr1, gc2,gr2, styles[cbWest], styles[cbNorth],
          styles[cbEast], styles[cbSouth], NO_CELL_BORDER, NO_CELL_BORDER);
        // ... Restores base border style overwritten in prev instruction
        Worksheet.WriteBorderStyles(cell, saved_styles);
        Worksheet.WriteBorders(cell, AValue);
      end else
      begin
        Worksheet.WriteBorders(cell, AValue);
        FixNeighborCellBorders(cell);
      end;

    finally
      EndUpdate;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBorders(
  ALeft, ATop, ARight, ABottom: Integer; AValue: TsCellBorders);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetCellBorder(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBorderStyle(ACol, ARow: Integer;
  ABorder: TsCellBorder; AValue: TsCellBorderStyle);
var
  cell: PCell;
  borders: TsCellBorders;
begin
  if Assigned(Worksheet) then begin
    BeginUpdate;
    try
      cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
      if Worksheet.IsMergeBase(cell) then
      begin
        borders := Worksheet.ReadCellBorders(cell);
        Worksheet.WriteBorderStyle(cell, ABorder, AValue);
        // This will apply the new border style to the outer cells of the range.
        SetCellBorder(ACol, ARow, borders);
      end else
      begin
        Worksheet.WriteBorderStyle(cell, ABorder, AValue);
        FixNeighborCellBorders(cell);
      end;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellBorderStyles(ALeft, ATop,
  ARight, ABottom: Integer; ABorder: TsCellBorder; AValue: TsCellBorderStyle);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetCellBorderStyle(c, r, ABorder, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellComment(ACol, ARow: Integer;
  AValue: String);
begin
  if Assigned(Worksheet) then
    Worksheet.WriteComment(GetWorksheetRow(ARow), GetWorksheetCol(ACol), AValue);
end;

procedure TsCustomWorksheetGrid.SetCellFont(ACol, ARow: Integer; AValue: TFont);
var
  fnt: TsFont;
  cell: PCell;
begin
  FCellFont.Assign(AValue);
  if Assigned(Worksheet) then begin
    fnt := TsFont.Create;
    try
      Convert_Font_To_sFont(FCellFont, fnt);
      cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
      Worksheet.WriteFont(cell, fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
    finally
      fnt.Free;
    end;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFonts(ALeft, ATop, ARight, ABottom: Integer;
  AValue: TFont);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetCellFont(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontColor(ACol, ARow: Integer; AValue: TsColor);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteFontColor(cell, AValue);
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontColors(
  ALeft, ATop, ARight, ABottom: Integer; AValue: TsColor);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetCellFontColor(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontName(ACol, ARow: Integer; AValue: String);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteFontName(cell, AValue);
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontNames(
  ALeft, ATop, ARight, ABottom: Integer; AValue: String);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetCellFontName(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontSize(ACol, ARow: Integer;
  AValue: Single);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteFontSize(cell, AValue);
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontSizes(
  ALeft, ATop, ARight, ABottom: Integer; AValue: Single);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetCellFontSize(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontStyle(ACol, ARow: Integer;
  AValue: TsFontStyles);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteFontStyle(cell, AValue);
  end;
end;

procedure TsCustomWorksheetGrid.SetCellFontStyles(
  ALeft, ATop, ARight, ABottom: Integer; AValue: TsFontStyles);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetCellFontStyle(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetCellValue(ACol, ARow: Integer; AValue: Variant);
var
  cell: PCell = nil;
  fmt: PsCellFormat = nil;
  nfp: TsNumFormatParams;
  r, c: Cardinal;
  s, plain: String;
  rtParams: TsRichTextParams;
begin
  if not Assigned(Worksheet) then
    exit;

  r := GetWorksheetRow(ARow);
  c := GetWorksheetCol(ACol);

  // If the cell already exists and contains a formula then the formula must be
  // removed. The formula would dominate over the data value.
  cell := Worksheet.FindCell(r, c);
  if HasFormula(cell) then cell^.FormulaValue := '';

  if VarIsNull(AValue) then
    Worksheet.WriteBlank(r, c)
  else
  if VarIsStr(AValue) then
  begin
    s := VarToStr(AValue);
    if (s <> '') and (s[1] = '=') then
      Worksheet.WriteFormula(r, c, Copy(s, 2, Length(s)), true)
    else
    begin
      cell := Worksheet.GetCell(r, c);
      HTMLToRichText(Workbook, Worksheet.ReadCellFont(cell), s, plain, rtParams);
      Worksheet.WriteText(cell, plain, rtParams);  // This will erase a non-formatted cell if s = ''
    end;
  end else
  if VarIsType(AValue, varDate) then
    Worksheet.WriteDateTime(r, c, VarToDateTime(AValue))
  else
  if VarIsNumeric(AValue) then
  begin
    // Check if the cell already exists and contains a format.
    // If it is a date/time format write a date/time cell...
    if cell <> nil then
    begin
//      fmt := Workbook.GetPointerToCellFormat(cell^.FormatIndex);
      fmt := Worksheet.GetPointerToEffectiveCellFormat(cell);
      if fmt <> nil then
        nfp := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);
      if (fmt <> nil) and IsDateTimeFormat(nfp) then
        Worksheet.WriteDateTime(r, c, VarToDateTime(AValue)) else
        Worksheet.WriteNumber(r, c, AValue);
    end
    else
      // ... otherwise write a number cell
      Worksheet.WriteNumber(r, c, AValue);
  end else
  if VarIsBool(AValue) then
    Worksheet.WriteBoolValue(r, c, AValue);
end;

procedure TsCustomWorksheetGrid.SetColWidths(ACol: Integer; AValue: Integer);
begin
  if GetColWidths(ACol) = AValue then
    exit;
  inherited ColWidths[ACol] := AValue;
  HeaderSized(true, ACol);
end;

procedure TsCustomWorksheetGrid.SetDefColWidth(AValue: Integer);
begin
  if (AValue = GetDefColWidth) or (AValue < 0) then
    exit;

  { AValue contains the zoom factor.
    FDefColWidth1000 is the col width at zoom factor 1.0 }
  FDefColWidth100 := round(AValue / ZoomFactor);
  inherited DefaultColWidth := AValue;
  if (FHeaderCount > 0) and HandleAllocated then begin
    PrepareCanvasFont;
    ColWidths[0] := GetDefaultHeaderColWidth;
  end;
  if (FZoomLock = 0) and (Worksheet <> nil) then
    Worksheet.WriteDefaultColWidth(CalcWorksheetColWidth(GetDefColWidth), Workbook.Units);
end;

procedure TsCustomWorksheetGrid.SetDefRowHeight(AValue: Integer);
begin
  if (AValue = GetDefRowHeight) or (AValue < 0) then
    exit;
  { AValue contains the zoom factor
    FDefRowHeight100 is the row height at zoom factor 1.0 }
  FDefRowHeight100 := round(AValue / ZoomFactor);
  inherited DefaultRowHeight := AValue;
  if FHeaderCount > 0 then
    RowHeights[0] := GetDefaultRowHeight;
  if (FZoomLock = 0) and (Worksheet <> nil) then
    Worksheet.WriteDefaultRowHeight(CalcWorksheetRowHeight(GetDefaultRowHeight), Workbook.Units);
end;

procedure TscustomWorksheetGrid.SetEditorLineMode(AValue: TsEditorLineMode);
begin
  if FLineMode = AValue then
    exit;

  FLineMode := AValue;
  if (FLineMode = elmMultiline) and (FMultilineStringEditor = nil) then
  begin
    FMultilineStringEditor := TMultilineStringCellEditor.Create(nil);
    FMultilineStringEditor.name :='StringEditor';
    FMultilineStringEditor.Text:='';
    FMultilineStringEditor.Visible:=False;
    FMultilineStringEditor.Align:=alNone;
    FMultilineStringEditor.BorderStyle := bsNone;
  end;
end;

procedure TsCustomWorksheetGrid.SetFrozenBorderPen(AValue: TPen);
begin
  FFrozenBorderPen.Assign(AValue);
  InvalidateGrid;
end;

procedure TsCustomWorksheetGrid.SetFrozenCols(AValue: Integer);
begin
  FFrozenCols := AValue;
  if Worksheet <> nil then begin
    Worksheet.LeftPaneWidth := FFrozenCols;
    if (FFrozenCols > 0) or (FFrozenRows > 0) then
      Worksheet.Options := Worksheet.Options + [soHasFrozenPanes]
    else
      Worksheet.Options := Worksheet.Options - [soHasFrozenPanes];
  end;
  Setup;
end;

procedure TsCustomWorksheetGrid.SetFrozenRows(AValue: Integer);
begin
  FFrozenRows := AValue;
  if Worksheet <> nil then begin
    Worksheet.TopPaneHeight := FFrozenRows;
    if (FFrozenCols > 0) or (FFrozenRows > 0) then
      Worksheet.Options := Worksheet.Options + [soHasFrozenPanes]
    else
      Worksheet.Options := Worksheet.Options - [soHasFrozenPanes];
  end;
  Setup;
end;

procedure TsCustomWorksheetGrid.SetHorAlignment(ACol, ARow: Integer;
  AValue: TsHorAlignment);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteHorAlignment(cell, AValue);
  end;
end;

procedure TsCustomWorksheetGrid.SetHorAlignments(
  ALeft, ATop, ARight, ABottom: Integer; AValue: TsHorAlignment);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetHorAlignment(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetHyperlink(ACol, ARow: Integer;
  AValue: String);
var
  p: Integer;
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    if AValue <> '' then
    begin
      p := pos('|', AValue);
      if p > 0 then
        Worksheet.WriteHyperlink(cell, copy(AValue, 1, p-1), copy(AValue, p+1, MaxInt))
      else
        Worksheet.WriteHyperlink(cell, AValue);
    end else
      Worksheet.RemoveHyperlink(cell);
  end;
end;

procedure TsCustomWorksheetGrid.SetNumberFormat(ACol, ARow: Integer; AValue: String);
begin
  if Assigned(Worksheet) then
    Worksheet.WriteNumberFormat(GetWorksheetRow(ARow), GetWorksheetCol(ACol), nfCustom, AValue);
end;

procedure TsCustomWorksheetGrid.SetNumberFormats(
  ALeft, ATop, ARight, ABottom: Integer; AValue: String);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetNumberFormat(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetReadFormulas(AValue: Boolean);
var
  optns: TsWorkbookOptions;
begin
  FReadFormulas := AValue;
  if Assigned(WorkbookSource) then
  begin
    optns := WorkbookSource.Options;
    if FReadFormulas then
      Include(optns, boReadFormulas)
    else
      Exclude(optns, boReadFormulas);
    WorkbookSource.Options := optns;
    if FInternalWorkbookSource <> nil then
      FInternalWorkbookSource.Options := optns;
  end;
end;

procedure TsCustomWorksheetGrid.SetRowHeights(ARow: Integer; AValue: Integer);
begin
  if GetRowHeights(ARow) = AValue then
    exit;
  inherited RowHeights[ARow] := AValue;
  HeaderSized(false, ARow);
end;

procedure TsCustomWorksheetGrid.SetSelPen(AValue: TsSelPen);
begin
  FSelPen.Assign(AValue);
  InvalidateGrid;
end;

{ Shows / hides the worksheet's grid lines }
procedure TsCustomWorksheetGrid.SetShowGridLines(AValue: Boolean);
begin
  if AValue = GetShowGridLines then
    Exit;

  if AValue then
    Options := Options + [goHorzLine, goVertLine]
  else
    Options := Options - [goHorzLine, goVertLine];

  if Worksheet <> nil then
    if AValue then
      Worksheet.Options := Worksheet.Options + [soShowGridLines]
    else
      Worksheet.Options := Worksheet.Options - [soShowGridLines];
end;

{ Shows / hides the worksheet's row and column headers. }
procedure TsCustomWorksheetGrid.SetShowHeaders(AValue: Boolean);
var
  hdrCount: Integer;
begin
  if AValue = GetShowHeaders then Exit;

  // Avoid crash if selected cell is at 0/0
  hdrCount := ord(AValue);
  if hdrCount > 0 then
  begin
    if Col < hdrCount then Col := hdrCount;
    if Row < hdrCount then Row := hdrCount;
  end;

  FHeaderCount := hdrCount;

  if Worksheet <> nil then
    if AValue then
      Worksheet.Options := Worksheet.Options + [soShowHeaders]
    else
      Worksheet.Options := Worksheet.Options - [soShowHeaders];

  Setup;
end;

procedure TsCustomWorksheetGrid.SetTextRotation(ACol, ARow: Integer;
  AValue: TsTextRotation);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteTextRotation(cell, AValue);
  end;
end;

procedure TsCustomWorksheetGrid.SetTextRotations(
  ALeft, ATop, ARight, ABottom: Integer; AValue: TsTextRotation);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetTextRotation(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetVertAlignment(ACol, ARow: Integer;
  AValue: TsVertAlignment);
var
  cell: PCell;
begin
  if Assigned(Worksheet) then
  begin
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteVertAlignment(cell, AValue);
  end;
end;

procedure TsCustomWorksheetGrid.SetVertAlignments(
  ALeft, ATop, ARight, ABottom: Integer; AValue: TsVertAlignment);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetVertAlignment(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetWordwrap(ACol, ARow: Integer;
  AValue: Boolean);
var
  cell: PCell;
begin
  if not Assigned(Worksheet) then
    exit;

  BeginUpdate;
  try
    cell := Worksheet.GetCell(GetWorksheetRow(ARow), GetWorksheetCol(ACol));
    Worksheet.WriteWordwrap(cell, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetWordwraps(ALeft, ATop, ARight, ABottom: Integer;
  AValue: Boolean);
var
  c,r: Integer;
begin
  EnsureOrder(ALeft, ARight);
  EnsureOrder(ATop, ABottom);
  BeginUpdate;
  try
    for c := ALeft to ARight do
      for r := ATop to ABottom do
        SetWordwrap(c, r, AValue);
  finally
    EndUpdate;
  end;
end;

procedure TsCustomWorksheetGrid.SetZoomFactor(AValue: Double);
begin
  if (AValue <> GetZoomFactor) and Assigned(Worksheet) then begin
    BeginUpdate;
    try
      Worksheet.ZoomFactor := abs(AValue);
  //    AdaptToZoomFactor;
    finally
      EndUpdate;
    end;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Registers the worksheet grid in the Lazarus component palette,
  page "FPSpreadsheet".
-------------------------------------------------------------------------------}
procedure Register;
begin
  RegisterComponents('FPSpreadsheet', [TsWorksheetGrid]);
end;


initialization
  fpsutils.ScreenPixelsPerInch := Screen.PixelsPerInch;
  FillPatternStyle := fsNoFill;

  RegisterPropertyToSkip(TsCustomWorksheetGrid, 'ColWidths',  'taken from worksheet', '');
  RegisterPropertyToSkip(TsCustomWorksheetGrid, 'RowHeights', 'taken from worksheet', '');

  crDragCopy := 1; //201705;
  Screen.Cursors[crDragCopy] := LoadCursorFromLazarusResource('cur_dragcopy');


finalization
  FreeAndNil(FillPatternBitmap);
  FreeAndNil(DragborderBitmap);

end.
