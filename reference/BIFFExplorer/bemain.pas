unit beMain;

{$mode objfpc}{$H+}

interface

uses
  ActnList, Classes, ComCtrls, ExtCtrls, Grids, Menus, StdCtrls, SysUtils,
  FileUtil, Forms, Controls, Graphics, Dialogs, Buttons, Types, VirtualTrees,
  {$ifdef USE_NEW_OLE}
  fpolebasic,
  {$else}
  fpolestorage,
  {$endif}
  fpstypes, KHexEditor,
  mrumanager, beTypes, beBIFFGrid;

type

  { TMainForm }
  TMainForm = class(TForm)
    AcFileOpen: TAction;
    AcFileQuit: TAction;
    AcFind: TAction;
    AcFindNext: TAction;
    AcFindPrev: TAction;
    AcAbout: TAction;
    AcFindClose: TAction;
    AcNodeExpand: TAction;
    AcNodeCollapse: TAction;
    AcDumpToFile: TAction;
    ActionList: TActionList;
    BIFFTree: TVirtualStringTree;
    CbFind: TComboBox;
    CbHexAddress: TCheckBox;
    CbHexEditorLineSize: TComboBox;
    CbHexSingleBytes: TCheckBox;
    ImageList: TImageList;
    HexEditor: TKHexEditor;
    MainMenu: TMainMenu;
    AnalysisDetails: TMemo;
    MenuItem1: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MnuDumpToFile: TMenuItem;
    MenuItem9: TMenuItem;
    MnuFind: TMenuItem;
    MnuRecord: TMenuItem;
    MnuFileReopen: TMenuItem;
    MenuItem4: TMenuItem;
    MnuHelp: TMenuItem;
    MenuItem2: TMenuItem;
    MnuFileQuit: TMenuItem;
    MnuFileOpen: TMenuItem;
    MnuFile: TMenuItem;
    OpenDialog: TOpenDialog;
    PageControl: TPageControl;
    DetailPanel: TPanel;
    FindPanel: TPanel;
    HexEditorParamsPanel: TPanel;
    SaveDialog: TSaveDialog;
    SpeedButton3: TSpeedButton;
    TreePopupMenu: TPopupMenu;
    TreePanel: TPanel;
    BtnFindNext: TSpeedButton;
    BtnFindPrev: TSpeedButton;
    RecentFilesPopupMenu: TPopupMenu;
    BtnCloseFind: TSpeedButton;
    Splitter1: TSplitter;
    HexSplitter: TSplitter;
    PgAnalysis: TTabSheet;
    PgValues: TTabSheet;
    DetailsSplitter: TSplitter;
    StatusBar: TStatusBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ValueGrid: TStringGrid;
    ToolBar: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    procedure AcAboutExecute(Sender: TObject);
    procedure AcDumpToFileExecute(Sender: TObject);
    procedure AcFileOpenExecute(Sender: TObject);
    procedure AcFileQuitExecute(Sender: TObject);
    procedure AcFindCloseExecute(Sender: TObject);
    procedure AcFindExecute(Sender: TObject);
    procedure AcFindNextExecute(Sender: TObject);
    procedure AcFindPrevExecute(Sender: TObject);
    procedure AcNodeCollapseExecute(Sender: TObject);
    procedure AcNodeCollapseUpdate(Sender: TObject);
    procedure AcNodeExpandExecute(Sender: TObject);
    procedure AcNodeExpandUpdate(Sender: TObject);
    procedure BIFFTreeBeforeCellPaint(Sender: TBaseVirtualTree;
      TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
    procedure BIFFTreeFocusChanged(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Column: TColumnIndex);
    procedure BIFFTreeFreeNode(Sender: TBaseVirtualTree; Node: PVirtualNode);
    procedure BIFFTreeGetNodeDataSize(Sender: TBaseVirtualTree;
      var NodeDataSize: Integer);
    procedure BIFFTreeGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: String);
//    procedure BIFFTreeInitNode(Sender: TBaseVirtualTree; ParentNode,
//      Node: PVirtualNode; var InitialStates: TVirtualNodeInitStates);
    procedure BIFFTreePaintText(Sender: TBaseVirtualTree;
      const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
      TextType: TVSTTextType);
    procedure CbFindChange(Sender: TObject);
    procedure CbFindKeyPress(Sender: TObject; var Key: char);
    procedure CbHexAddressChange(Sender: TObject);
    procedure CbHexEditorLineSizeChange(Sender: TObject);
    procedure CbHexSingleBytesChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormDropFiles(Sender: TObject; const FileNames: array of String);
    procedure FormShow(Sender: TObject);
    procedure HexEditorClick(Sender: TObject);
    procedure HexEditorKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ListViewSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure PageControlChange(Sender: TObject);
    procedure ValueGridClick(Sender: TObject);
    procedure ValueGridPrepareCanvas(Sender: TObject; ACol, ARow: Integer;
      AState: TGridDrawState);

  private
    MemStream: TMemoryStream;
    OLEStorage: TOLEStorage;
    FFileName: String;
    FFormat: TsSpreadsheetFormat;
    FBuffer: TBIFFBuffer;
    FCurrOffset: Integer;
    FXFIndex: Integer;
    FFontIndex: Integer;
    FFormatIndex: Integer;
    FRowIndex: Integer;
    FExternSheetIndex: Integer;
    FAnalysisGrid: TBIFFGrid;
    FMRUMenuManager : TMRUMenuManager;
    procedure AddToHistory(const AText: String);
    procedure AnalysisGridDetails(Sender: TObject; ADetails: TStrings);
    procedure AnalysisGridPrepareCanvas(Sender: TObject; ACol, ARow: Integer;
      AState: TGridDrawState);
    procedure AnalysisGridSelection(Sender: TObject; ACol, ARow: Integer);
    procedure DumpToFile(const AFileName: String);
    procedure ExecFind(ANext, AKeep: Boolean);
    function  GetBIFFNodeData: PBiffNodeData;
    function  GetRecType: Word;
    function  GetValueGridDataSize: Integer;
    procedure LoadFile(const AFileName: String); overload;
    procedure LoadFile(const AFileName: String; AFormat: TsSpreadsheetFormat); overload;
    procedure MRUMenuManagerRecentFile(Sender:TObject; const AFileName:string);
    procedure PopulateAnalysisGrid;
    procedure PopulateHexDump;
    procedure PopulateValueGrid;
    procedure ReadCmdLine;
    procedure ReadFromIni;
    procedure ReadFromStream(AStream: TStream);
    procedure UpdateCaption;
    procedure UpdateCmds;
    procedure UpdateStatusbar;
    procedure WriteToIni;

  public
    procedure BeforeRun;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

uses
  IniFiles, LazUTF8, LazFileUtils, Math, StrUtils, LCLType,
  KFunctions,
  fpsUtils,
  beUtils, beBIFFUtils, beAbout;

const
  VALUE_ROW_INDEX      = 1;
  VALUE_ROW_BITS       = 2;
  VALUE_ROW_BYTE       = 3;
  VALUE_ROW_SHORTINT   = 4;
  VALUE_ROW_WORD       = 5;
  VALUE_ROW_SMALLINT   = 6;
  VALUE_ROW_DWORD      = 7;
  VALUE_ROW_LONGINT    = 8;
  VALUE_ROW_QWORD      = 9;
  VALUE_ROW_INT64      = 10;
  VALUE_ROW_SINGLE     = 11;
  VALUE_ROW_DOUBLE     = 12;
  VALUE_ROW_ANSISTRING = 13;
  VALUE_ROW_PANSICHAR  = 14;
  VALUE_ROW_WIDESTRING = 15;
  VALUE_ROW_PWIDECHAR  = 16;

  MAX_HISTORY = 16;


{ TMyHexEditor }

type
  TMyHexEditor = class(TKHexEditor);

       (*
{ Virtual tree nodes }

type
  TObjectNodeData = record
    Data: TObject;
  end;
  PObjectNodeData = ^TObjectNodeData;
         *)

{ TMainForm }

procedure TMainForm.AcAboutExecute(Sender: TObject);
var
  F: TAboutForm;
begin
  F := TAboutForm.Create(nil);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;


procedure TMainForm.AcDumpToFileExecute(Sender: TObject);
begin
  if FFileName = '' then
    exit;

  with SaveDialog do begin
    FileName := ChangeFileExt(ExtractFileName(FFileName), '') + '_dumped.txt';
    if Execute then
      DumpToFile(FileName);
  end;
end;


procedure TMainForm.AcFileOpenExecute(Sender: TObject);
begin
  with OpenDialog do begin
    if Execute then LoadFile(FileName);
  end;
end;


procedure TMainForm.AcFileQuitExecute(Sender: TObject);
begin
  Close;
end;


procedure TMainForm.AcFindCloseExecute(Sender: TObject);
begin
  AcFind.Checked := false;
  FindPanel.Hide;
end;


procedure TMainForm.AcFindExecute(Sender: TObject);
begin
  if AcFind.Checked then begin
    FindPanel.Show;
    CbFind.SetFocus;
  end else begin
    FindPanel.Hide;
  end;
end;


procedure TMainForm.AcFindNextExecute(Sender: TObject);
begin
  ExecFind(true, false);
end;


procedure TMainForm.AcFindPrevExecute(Sender: TObject);
begin
  ExecFind(false, false);
end;


procedure TMainForm.AcNodeCollapseExecute(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
    BiffTree.Expanded[node] := false;
  end;
end;

procedure TMainForm.AcNodeCollapseUpdate(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
   end;
  AcNodeCollapse.Enabled := (node <> nil) and BiffTree.Expanded[node];
end;


procedure TMainForm.AcNodeExpandExecute(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
    BiffTree.Expanded[node] := true;
  end;
end;


procedure TMainForm.AcNodeExpandUpdate(Sender: TObject);
var
  node: PVirtualNode;
begin
  node := BiffTree.FocusedNode;
  if node <> nil then begin
    if BiffTree.GetNodeLevel(node) > 0 then
      node := node^.Parent;
  end;
  AcNodeExpand.Enabled := (node <> nil) and not BiffTree.Expanded[node];
end;

procedure TMainForm.AddToHistory(const AText: String);
begin
  if (AText <> '') and (CbFind.Items.IndexOf(AText) = -1) then begin
    CbFind.Items.Insert(0, AText);
    while CbFind.Items.Count > MAX_HISTORY do
      CbFind.Items.Delete(CbFind.Items.Count-1);
  end;
end;


procedure TMainForm.AnalysisGridDetails(Sender: TObject; ADetails: TStrings);
begin
  AnalysisDetails.Lines.Assign(ADetails);
end;


procedure TMainForm.AnalysisGridPrepareCanvas(Sender: TObject; ACol,
  ARow: Integer; AState: TGridDrawState);
begin
  if ARow = 0 then FAnalysisGrid.Canvas.Font.Style := [fsBold];
end;


procedure TMainForm.AnalysisGridSelection(Sender: TObject; ACol, ARow: Integer);
var
  s: String;
begin
  if ARow < FAnalysisGrid.RowCount then
  begin
    s := FAnalysisGrid.Cells[0, ARow];
    if s <> '' then
    begin
      FCurrOffset := StrToInt(s);
      PopulateValueGrid;
      UpdateStatusbar;
    end;
  end;
end;


procedure TMainForm.BeforeRun;
begin
  ReadFromIni;
  ReadCmdLine;
end;


procedure TMainForm.BIFFTreeBeforeCellPaint(Sender: TBaseVirtualTree;
  TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  CellPaintMode: TVTCellPaintMode; CellRect: TRect; var ContentRect: TRect);
var
  s: String;
begin
  if (Sender.GetNodeLevel(Node) = 0) and (Column = 0) then begin
    // Left-align parent nodes (column 0 is right-aligned)
    BiffTreeGetText(Sender, Node, 0, ttNormal, s);
    TargetCanvas.Font.Style := [fsBold];
    ContentRect.Right := CellRect.Left + TargetCanvas.TextWidth(s) + 30;
  end;
end;


procedure TMainForm.BIFFTreeFocusChanged(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex);
var
  data: PBiffNodeData;
  n: Word;
begin
  if Node^.Parent = Sender.RootNode then
  begin
    HexEditor.Clear;
    for n:=1 to ValueGrid.RowCount-1 do
    begin
      ValueGrid.Cells[1, n] := '';
      ValueGrid.Cells[2, n] := '';
    end;
    FAnalysisGrid.RowCount := 2;
    FAnalysisGrid.Rows[1].Clear;
    AnalysisDetails.Lines.Clear;
    exit;
  end;

  data := Sender.GetNodeData(Node);

  // Move to start of record + 2 bytes to skip record type ID.
  MemStream.Position := PtrInt(data^.Offset) + 2;

  // Read size of record
  n := WordLEToN(MemStream.ReadWord);

  // Read record data
  SetLength(FBuffer, n);
  if n > 0 then
    MemStream.ReadBuffer(FBuffer[0], n);

  // Update user interface
  if (BiffTree.FocusedNode <> nil) and (BiffTree.GetNodeLevel(BiffTree.FocusedNode) > 0)
  then begin
    Statusbar.Panels[0].Text := Format('Record ID: $%.4x', [data^.RecordID]);
    Statusbar.Panels[1].Text := data^.RecordName;
    Statusbar.Panels[2].Text := Format('Record size: %d bytes', [n]);
    Statusbar.Panels[3].Text := '';
  end else begin
    Statusbar.Panels[0].Text := '';
    Statusbar.Panels[1].Text := data^.RecordName;
    Statusbar.Panels[2].Text := '';
    Statusbar.Panels[3].Text := '';
  end;
  PopulateHexDump;
  PageControlChange(nil);
end;


procedure TMainForm.BIFFTreeFreeNode(Sender: TBaseVirtualTree;
  Node: PVirtualNode);
var
  data: PBiffNodeData;
begin
  data := Sender.GetNodeData(Node);
  if data <> nil then
  begin
    data^.RecordName := '';
    data^.RecordDescription := '';
  end;
end;


procedure TMainForm.BIFFTreeGetNodeDataSize(Sender: TBaseVirtualTree;
  var NodeDataSize: Integer);
begin
  NodeDataSize := SizeOf(TBiffNodeData);
end;


procedure TMainForm.BIFFTreeGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: String);
var
  data: PBiffNodeData;
begin
  CellText := '';
  data := Sender.GetNodeData(Node);
  if data <> nil then
    case Sender.GetNodeLevel(Node) of
      0: if Column = 0 then CellText := data^.RecordName;
      1: case Column of
           0: CellText := IntToStr(data^.Offset);
           1: CellText := Format('$%.4x', [data^.RecordID]);
           2: CellText := data^.RecordName;
           3: if data^.Index > -1 then CellText := IntToStr(data^.Index);
           4: cellText := data^.RecordDescription;
         end;
    end;
end;


procedure TMainForm.BIFFTreePaintText(Sender: TBaseVirtualTree;
  const TargetCanvas: TCanvas; Node: PVirtualNode; Column: TColumnIndex;
  TextType: TVSTTextType);
begin
  // Paint parent node in bold font.
  if (Sender.GetNodeLevel(Node) = 0) and (Column = 0) then
    TargetCanvas.Font.Style := [fsBold];
end;


procedure TMainForm.CbFindChange(Sender: TObject);
begin
  ExecFind(true, true);
end;


procedure TMainForm.CbFindKeyPress(Sender: TObject; var Key: char);
begin
  if Key = #13 then
    ExecFind(true, false);
end;

procedure TMainForm.CbHexAddressChange(Sender: TObject);
begin
  if CbHexAddress.Checked then
  begin
    HexEditor.AddressMode := eamHex;
    HexEditor.AddressPrefix := '$';
  end else
  begin
    HexEditor.AddressMode := eamDec;
    HexEditor.AddressPrefix := '';
  end;
  CbHexEditorLineSizeChange(nil);
end;

procedure TMainForm.CbHexEditorLineSizeChange(Sender: TObject);
begin
  case CbHexEditorLineSize.ItemIndex of
    0: HexEditor.LineSize := IfThen(HexEditor.AddressMode = eamHex, 16, 10);
    1: HexEditor.LineSize := IfThen(HexEditor.AddressMode = eamHex, 32, 20);
  end;
end;

procedure TMainForm.CbHexSingleBytesChange(Sender: TObject);
begin
  HexEditor.DigitGrouping := IfThen(CbHexSingleBytes.Checked, 1, 2);
end;

procedure TMainForm.DumpToFile(const AFileName: String);
var
  list: TStringList;
  parentnode, node: PVirtualNode;
  parentdata, data: PBiffNodeData;
begin
  list := TStringList.Create;
  try
    parentnode := BiffTree.GetFirst;
    while parentnode <> nil do begin
      parentdata := BiffTree.GetNodeData(parentnode);
      list.Add(parentdata^.RecordName);
      node := BIffTree.GetFirstChild(parentnode);
      while node <> nil do begin
        data := BiffTree.GetNodeData(node);
        List.Add(Format('  %.04x %s (%s)', [data^.RecordID, data^.RecordName, data^.RecordDescription]));
        node := BiffTree.GetNextSibling(node);
      end;
      List.Add('');
      parentnode := BiffTree.GetNextSibling(parentnode);
    end;

    list.SaveToFile(AFileName);
  finally
    list.Free;
  end;
end;


procedure TMainForm.ExecFind(ANext, AKeep: Boolean);
var
  s: String;
  node, node0: PVirtualNode;

  function GetRecordname(ANode: PVirtualNode; UseLowercase: Boolean = true): String;
  var
    data: PBIffNodeData;
  begin
    data := BiffTree.GetNodeData(ANode);
    if Assigned(data) then begin
      if UseLowercase then
        Result := lowercase(data^.RecordName)
      else
        Result := data^.RecordName;
    end else
      Result := '';
  end;

  function GetNextNode(ANode: PVirtualNode): PVirtualNode;
  var
    nextparent: PVirtualNode;
  begin
    Result := BiffTree.GetNextSibling(ANode);
    if (Result = nil) and (ANode <> nil) then begin
      nextparent := BiffTree.GetNextSibling(ANode^.Parent);
      if nextparent = nil then
        nextparent := BiffTree.GetFirst;
      Result := BiffTree.GetFirstChild(nextparent);
    end;
  end;

  function GetPrevNode(ANode: PVirtualNode): PVirtualNode;
  var
    prevparent: PVirtualNode;
  begin
    Result := BiffTree.GetPreviousSibling(ANode);
    if (Result = nil) and (ANode <> nil) then begin
      prevparent := BiffTree.GetPreviousSibling(ANode^.Parent);
      if prevparent = nil then
        prevparent := BiffTree.GetLast;
      Result := BiffTree.GetLastChild(prevparent);
    end;
  end;

begin
  if CbFind.Text = '' then
    exit;

  s := Lowercase(CbFind.Text);
  node0 := BiffTree.FocusedNode;
  if node0 = nil then
    node0 := BiffTree.GetFirst;
  if BiffTree.GetNodeLevel(node0) = 0 then
    node0 := BiffTree.GetFirstChild(node0);

  if ANext then begin
    if AKeep
      then node := node0
      else node := GetNextNode(node0);
    repeat
      if pos(s, GetRecordname(node)) > 0 then begin
        BiffTree.FocusedNode := node;
        BiffTree.Selected[node] := true;
        BiffTree.ScrollIntoView(node, true);
        AddToHistory(GetRecordname(node, false));
        exit;
      end;
      node := GetNextNode(node);
    until (node = node0) or (node = nil);
  end else begin
    if AKeep
      then node := node0
      else node := GetPrevNode(node0);
    repeat
      if pos(s, GetRecordName(node)) > 0 then begin
        BiffTree.FocusedNode := node;
        BiffTree.Selected[node] := true;
        BiffTree.ScrollIntoView(node, true);
        AddToHistory(GetRecordName(node, false));
        exit;
      end;
      node := GetPrevNode(node);
    until (node = node0) or (node = nil);
  end;
end;


procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if CanClose then
    try
      WriteToIni;
    except
      MessageDlg('Could not write setting to configuration file.', mtError, [mbOK], 0);
    end;
end;


procedure TMainForm.FormCreate(Sender: TObject);
begin
  FMRUMenuManager := TMRUMenuManager.Create(self);
  with FMRUMenuManager do begin
    Name := 'MRUMenuManager';
    IniFileName := GetAppConfigFile(false);
    IniSection := 'RecentFiles';
    MaxRecent := 16;
    MenuCaptionMask := '&%x - %s';    // & --> create hotkey
    MenuItem := MnuFileReopen;
    PopupMenu := RecentFilesPopupMenu;
    OnRecentFile := @MRUMenuManagerRecentFile;
  end;

  HexEditor.Font.Style := [];

  FAnalysisGrid := TBIFFGrid.Create(self);
  with FAnalysisGrid do begin
    Parent := PgAnalysis;
    Align := alClient;
    DefaultRowHeight := ValueGrid.DefaultRowHeight;
    Options := Options + [goDrawFocusSelected];
    TitleStyle := tsNative;
    OnDetails := @AnalysisGridDetails;
    OnPrepareCanvas := @AnalysisGridPrepareCanvas;
    OnSelection := @AnalysisGridSelection;
    TabOrder := 0;
  end;

  with ValueGrid do begin
    ColCount := 3;
    RowCount := VALUE_ROW_PWIDECHAR + 1;
    Cells[0, 0] := 'Data type';
    Cells[1, 0] := 'Value';
    Cells[2, 0] := 'Offset range';
    Cells[0, VALUE_ROW_INDEX] := 'Offset';
    Cells[0, VALUE_ROW_BITS] := 'Bits';
    Cells[0, VALUE_ROW_BYTE] := 'Byte';
    Cells[0, VALUE_ROW_SHORTINT] := 'ShortInt';
    Cells[0, VALUE_ROW_WORD] := 'Word';
    Cells[0, VALUE_ROW_SMALLINT] := 'SmallInt';
    Cells[0, VALUE_ROW_DWORD] := 'DWord';
    Cells[0, VALUE_ROW_LONGINT] := 'LongInt';
    Cells[0, VALUE_ROW_QWORD] := 'QWord';
    Cells[0, VALUE_ROW_INT64] := 'Int64';
    Cells[0, VALUE_ROW_SINGLE] := 'Single';
    Cells[0, VALUE_ROW_DOUBLE] := 'Double';
    Cells[0, VALUE_ROW_ANSISTRING] := 'AnsiString';
    Cells[0, VALUE_ROW_PANSICHAR] := 'PAnsiChar';
    Cells[0, VALUE_ROW_WIDESTRING] := 'WideString';
    Cells[0, VALUE_ROW_PWIDECHAR] := 'PWideChar';
  end;

  BiffTree.DefaultNodeHeight := BiffTree.Canvas.TextHeight('Tg') + 4;
  BiffTree.Header.DefaultHeight := ValueGrid.DefaultRowHeight;

  UpdateCmds;
end;


procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if MemStream <> nil then
    FreeAndNil(MemStream);
  if OLEStorage <> nil then
    FreeAndNil(OLEStorage);
end;


procedure TMainForm.FormDropFiles(Sender: TObject;
  const FileNames: array of String);
begin
  LoadFile(FileNames[0]);
end;


procedure TMainForm.FormShow(Sender: TObject);
begin
  Width := Width + 1;     // remove black rectangle next to ValueGrid
  Width := Width - 1;
end;


function TMainForm.GetBIFFNodeData: PBiffNodeData;
begin
  Result := nil;
  if BiffTree.FocusedNode <> nil then
  begin
    Result := BiffTree.GetNodeData(BiffTree.FocusedNode);
    if Result <> nil then
      MemStream.Position := Result^.Offset;
  end;
end;


function TMainForm.GetRecType: Word;
var
  data: PBiffNodeData;
begin
  Result := Word(-1);
  if BiffTree.FocusedNode <> nil then
  begin
    data := BiffTree.GetNodedata(BiffTree.FocusedNode);
    if data <> nil then
    begin
      MemStream.Position := data^.Offset;
      Result := WordLEToN(MemStream.ReadWord);
    end;
  end;
end;


function TMainForm.GetValueGridDataSize: Integer;
begin
  Result := -1;
  case ValueGrid.Row of
    VALUE_ROW_BITS     : Result := SizeOf(Byte);
    VALUE_ROW_BYTE     : Result := SizeOf(Byte);
    VALUE_ROW_SHORTINT : Result := SizeOf(ShortInt);
    VALUE_ROW_WORD     : Result := SizeOf(Word);
    VALUE_ROW_SMALLINT : Result := SizeOf(SmallInt);
    VALUE_ROW_DWORD    : Result := SizeOf(DWord);
    VALUE_ROW_LONGINT  : Result := SizeOf(LongInt);
    VALUE_ROW_QWORD    : Result := SizeOf(QWord);
    VALUE_ROW_INT64    : Result := SizeOf(Int64);
    VALUE_ROW_SINGLE   : Result := SizeOf(Single);
    VALUE_ROW_DOUBLE   : Result := SizeOf(Double);
  end;
end;

procedure TMainForm.HexEditorClick(Sender: TObject);
begin
  FCurrOffset := HexEditor.SelStart.Index;
  PopulateValueGrid;
  ValueGridClick(nil);
  UpdateStatusbar;
end;

procedure TMainForm.HexEditorKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  sel: TKHexEditorSelection;
begin
  case Key of
    VK_LEFT  : dec(FCurrOffset);
    VK_RIGHT : inc(FCurrOffset);
    VK_UP    : dec(FCurrOffset, HexEditor.LineSize);
    VK_DOWN  : inc(FCurrOffset, HexEditor.LineSize);
    VK_HOME  : if (Shift = [ssCtrl]) then
                 FCurrOffset := 0 else
                 FCurrOffset := (FCurrOffset div HexEditor.LineSize) * HexEditor.LineSize;
    VK_END   : if (Shift = [ssCtrl]) then
                 FCurrOffset := High(FBuffer) else
                 FCurrOffset := succ(FCurrOffset div HexEditor.LineSize) * HexEditor.lineSize - 1;
    VK_NEXT  : begin
                 if (Shift = [ssCtrl]) then
                   inc(FCurrOffset, HexEditor.LineSize * HexEditor.LineCount)
                 else
                   inc(FCurrOffset, HexEditor.LineSize * HexEditor.GetClientHeightChars);
                 while (FCurrOffset > High(FBuffer)) do
                   dec(FCurrOffset, HexEditor.LineSize);
               end;
    VK_PRIOR : if (Shift = [ssCtrl]) then
                 FCurrOffset := FCurrOffset mod HexEditor.LineSize
               else
               begin
                 dec(FCurrOffset, HexEditor.LineSize * HexEditor.GetClientHeightChars);
                 while (FCurrOffset < 0) do
                   inc(FCurrOffset, HexEditor.LineSize);
               end;
    else       exit;
  end;
  if FCurrOffset < 0 then FCurrOffset := 0;
  if FCurrOffset > High(FBuffer) then FCurrOffset := High(FBuffer);
  sel.Index := FCurrOffset;
  sel.Digit := 0;
  HexEditor.SelStart := sel;
  HexEditorClick(nil);
  if not HexEditor.CaretInView then
    TMyHexEditor(HexEditor).ScrollTo(HexEditor.SelToPoint(HexEditor.SelStart, HexEditor.EditArea), false, true);

  // Don't process these keys any more!
  Key := 0;
end;

procedure TMainForm.LoadFile(const AFileName: String);
var
  valid: Boolean;
  excptn: Exception = nil;
  ext: String;
begin
  if not FileExistsUTF8(AFileName) then begin
    MessageDlg(Format('File "%s" not found.', [AFileName]), mtError, [mbOK], 0);
    exit;
  end;

  ext := Lowercase(ExtractFileExt(AFilename));
  if ext <> '.xls' then begin
    MessageDlg('BIFFExplorer can only process binary Excel files (extension ".xls")',
      mtError, [mbOK], 0);
    exit;
  end;

  // .xls files can contain several formats. We look into the header first.
  if ext = STR_EXCEL_EXTENSION then
  begin
    valid := GetFormatFromFileHeader(UTF8ToAnsi(AFileName), FFormat);
    // It is possible that valid xls files are not detected correctly. Therefore,
    // we open them explicitly by trial and error - see below.
    if not valid then
      FFormat := sfExcel8;
    valid := true;
  end else
    FFormat := sfExcel8;

  while True do begin
    try
      LoadFile(AFileName, FFormat);
      valid := True;
    except
      on E: Exception do begin
        if FFormat = sfExcel8 then excptn := E;
        valid := False
      end;
    end;
    if valid or (FFormat = sfExcel2) then Break;
    FFormat := Pred(FFormat);
  end;

  // A failed attempt to read a file should bring an exception, so re-raise
  // the exception if necessary. We re-raise the exception brought by Excel 8,
  // since this is the most common format
  if (not valid) and (excptn <> nil) then
    raise excptn;
end;


procedure TMainForm.LoadFile(const AFileName: String; AFormat: TsSpreadsheetFormat);
var
  OLEDocument: TOLEDocument;
  streamname: UTF8String;
  filestream: TFileStream;
begin
  if MemStream <> nil then
    FreeAndNil(MemStream);

  if OLEStorage <> nil then
    FreeAndNil(OLEStorage);

  MemStream := TMemoryStream.Create;

  if AFormat = sfExcel2 then begin
    fileStream := TFileStream.Create(UTF8ToSys(AFileName), fmOpenRead + fmShareDenyNone);
    try
      MemStream.CopyFrom(fileStream, fileStream.Size);
      MemStream.Position := 0;
//    MemStream.LoadFromFile(UTF8ToSys(AFileName));
    finally
      filestream.Free;
    end;
  end else begin
    OLEStorage := TOLEStorage.Create;

    // Only one stream is necessary for any number of worksheets
    OLEDocument.Stream := MemStream;
    if AFormat = sfExcel8 then streamname := 'Workbook' else streamname := 'Book';
    OLEStorage.ReadOLEFile(UTF8ToSys(AFileName), OLEDocument, streamname);

    // Check if the operation succeded
    if MemStream.Size = 0 then
      raise Exception.Create('BIFF Explorer: Reading the OLE document failed');
  end;

  // Rewind the stream and read from it
  MemStream.Position := 0;
  FFileName := ExpandFileName(AFileName);
  ReadFromStream(MemStream);

  FFormat := AFormat;
  UpdateCaption;
  UpdateStatusbar;

  FMRUMenuManager.AddToRecent(AFileName);
end;


procedure TMainForm.ListViewSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  n: Word;
begin
  if Selected then begin
    // Move to start of record + 2 bytes to skip record type ID.
    MemStream.Position := PtrInt(Item.Data) + 2;

    // Read size of record
    n := WordLEToN(MemStream.ReadWord);

    // Read record data
    SetLength(FBuffer, n);
    MemStream.ReadBuffer(FBuffer[0], n);

    // Update user interface
    Statusbar.Panels[0].Text := Format('Record ID: %s', [Item.SubItems[0]]);
    Statusbar.Panels[1].Text := Item.SubItems[1];
    Statusbar.Panels[2].Text := Format('Record size: %s bytes', [Item.SubItems[3]]);
    PopulateHexDump;
    PageControlChange(nil);
  end;
end;


procedure TMainForm.MRUMenuManagerRecentFile(Sender: TObject;
  const AFileName: string);
begin
  LoadFile(AFileName);
end;


procedure TMainForm.PopulateAnalysisGrid;
begin
  FAnalysisGrid.SetBIFFNodeData(GetBiffNodeData, FBuffer, FFormat);
end;


procedure TMainForm.PopulateHexDump;
var
  data: TDataSize;
begin
  data.Size := Length(FBuffer);
  data.Data := @FBuffer[0];
  HexEditor.Clear;
  HexEditor.Append(0, data);
end;


procedure TMainForm.PopulateValueGrid;
var
  buf: array[0..1023] of Byte;
  w: word absolute buf;
  dw: DWord absolute buf;
  qw: QWord absolute buf;
  dbl: double absolute buf;
  sng: single absolute buf;
  idx: Integer;
  i, j: Integer;
  s: String;
  sw: WideString;
  ls: Integer;
begin
  idx := FCurrOffset;
//  idx := HexEditor.SelStart.Index;

  i := ValueGrid.RowCount;
  j := ValueGrid.ColCount;

  ValueGrid.Cells[1, VALUE_ROW_INDEX] := IntToStr(idx);

  if idx <= Length(FBuffer)-SizeOf(byte) then begin
    ValueGrid.Cells[1, VALUE_ROW_BITS] := IntToBin(FBuffer[idx], 8);
    ValueGrid.Cells[2, VALUE_ROW_BITS] := Format('%d ... %d', [idx, idx]);
    ValueGrid.Cells[1, VALUE_ROW_BYTE] := IntToStr(FBuffer[idx]);
    ValueGrid.Cells[2, VALUE_ROW_BYTE] := ValueGrid.Cells[2, VALUE_ROW_BITS];
    ValueGrid.Cells[1, VALUE_ROW_SHORTINT] := IntToStr(ShortInt(FBuffer[idx]));
    ValueGrid.Cells[2, VALUE_ROW_SHORTINT] := ValueGrid.Cells[2, VALUE_ROW_BITS];
  end
  else begin
    ValueGrid.Cells[1, VALUE_ROW_BYTE] := '';
    ValueGrid.Cells[2, VALUE_ROW_BYTE] := '';
    ValueGrid.Cells[1, VALUE_ROW_SHORTINT] := '';
    ValueGrid.Cells[2, VALUE_ROW_SHORTINT] := '';
  end;

  if idx <= Length(FBuffer)-SizeOf(word) then begin
    buf[0] := FBuffer[idx];
    buf[1] := FBuffer[idx+1];
    ValueGrid.Cells[1, VALUE_ROW_WORD] := IntToStr(WordLEToN(w));
    ValueGrid.Cells[2, VALUE_ROW_WORD] := Format('%d ... %d', [idx, idx+SizeOf(Word)-1]);
    ValueGrid.Cells[1, VALUE_ROW_SMALLINT] := IntToStr(SmallInt(WordLEToN(w)));
    ValueGrid.Cells[2, VALUE_ROW_SMALLINT] := ValueGrid.Cells[2, VALUE_ROW_WORD];
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_WORD] := '';
    ValueGrid.Cells[2, VALUE_ROW_WORD] := '';
    ValueGrid.Cells[1, VALUE_ROW_SMALLINT] := '';
    ValueGrid.Cells[2, VALUE_ROW_SMALLINT] := '';
  end;

  if idx <= Length(FBuffer) - SizeOf(DWord) then begin
    for i:=0 to SizeOf(DWord)-1 do buf[i] := FBuffer[idx+i];
    ValueGrid.Cells[1, VALUE_ROW_DWORD] := IntToStr(DWordLEToN(dw));
    ValueGrid.Cells[2, VALUE_ROW_DWORD] := Format('%d ... %d', [idx, idx+SizeOf(DWord)-1]);
    ValueGrid.Cells[1, VALUE_ROW_LONGINT] := IntToStr(LongInt(DWordLEToN(dw)));
    ValueGrid.Cells[2, VALUE_ROW_LONGINT] := ValueGrid.Cells[2, VALUE_ROW_DWORD];
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_DWORD] := '';
    ValueGrid.Cells[2, VALUE_ROW_DWORD] := '';
    ValueGrid.Cells[1, VALUE_ROW_LONGINT] := '';
    ValueGrid.Cells[2, VALUE_ROW_LONGINT] := '';
  end;

  if idx <= Length(FBuffer) - SizeOf(QWord) then begin
    for i:=0 to SizeOf(QWord)-1 do buf[i] := FBuffer[idx+i];
    ValueGrid.Cells[1, VALUE_ROW_QWORD] := Format('%d', [qw]);
    ValueGrid.Cells[2, VALUE_ROW_QWORD] := Format('%d ... %d', [idx, idx+SizeOf(QWord)-1]);
    ValueGrid.Cells[1, VALUE_ROW_INT64] := Format('%d', [Int64(qw)]);
    ValueGrid.Cells[2, VALUE_ROW_INT64] := ValueGrid.Cells[2, VALUE_ROW_QWORD];
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_QWORD] := '';
    ValueGrid.Cells[2, VALUE_ROW_QWORD] := '';
    ValueGrid.Cells[1, VALUE_ROW_INT64] := '';
    ValueGrid.Cells[2, VALUE_ROW_INT64] := '';
  end;

  if idx <= Length(FBuffer) - SizeOf(single) then begin
    for i:=0 to SizeOf(single)-1 do buf[i] := FBuffer[idx+i];
    ValueGrid.Cells[1, VALUE_ROW_SINGLE] := Format('%f', [sng]);
    ValueGrid.Cells[2, VALUE_ROW_SINGLE] := Format('%d ... %d', [idx, idx+SizeOf(Single)-1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_SINGLE] := '';
    ValueGrid.Cells[2, VALUE_ROW_SINGLE] := '';
  end;

  if idx <= Length(FBuffer) - SizeOf(double) then begin
    for i:=0 to SizeOf(double)-1 do buf[i] := FBuffer[idx+i];
    ValueGrid.Cells[1, VALUE_ROW_DOUBLE] := Format('%f', [dbl]);
    ValueGrid.Cells[2, VALUE_ROW_DOUBLE] := Format('%d ... %d', [idx, idx+SizeOf(Double)-1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_DOUBLE] := '';
    ValueGrid.Cells[2, VALUE_ROW_DOUBLE] := '';
  end;

  if idx < Length(FBuffer) then begin
    ls := FBuffer[idx];
    SetLength(s, ls);
    i := idx + 1;
    j := 0;
    while (i < Length(FBuffer)) and (j < Length(s)) do begin
      inc(j);
      s[j] := char(FBuffer[i]);
      inc(i);
    end;
    SetLength(s, j);
    ValueGrid.Cells[1, VALUE_ROW_ANSISTRING] := s;
    ValueGrid.Cells[2, VALUE_ROW_ANSISTRING] := Format('%d ... %d', [idx, ls * SizeOf(char) + 1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_ANSISTRING] := '';
    ValueGrid.Cells[2, VALUE_ROW_ANSISTRING] := '';
  end;

  s := StrPas(PChar(@FBuffer[idx]));
  ValueGrid.Cells[1, VALUE_ROW_PANSICHAR] := s;
  ValueGrid.Cells[2, VALUE_ROW_PANSICHAR] := Format('%d ... %d', [idx, idx + Length(s)]);

  if idx < Length(FBuffer) then begin
    ls := FBuffer[idx];
    SetLength(sw, ls);
    j := 0;
    i := idx + 2;
    while (i < Length(FBuffer)-1) and (j < Length(sw)) do begin
      buf[0] := FBuffer[i];
      buf[1] := FBuffer[i+1];
      inc(i, SizeOf(WideChar));
      inc(j);
      sw[j] := WideChar(w);
    end;
    SetLength(sw, j);
    ValueGrid.Cells[1, VALUE_ROW_WIDESTRING] := UTF8Encode(sw);
    ValueGrid.Cells[2, VALUE_ROW_WIDESTRING] := Format('%d ... %d', [idx, idx + ls*SizeOf(wideChar)+1]);
  end else begin
    ValueGrid.Cells[1, VALUE_ROW_WIDESTRING] := '';
    ValueGrid.Cells[2, VALUE_ROW_WIDESTRING] := '';
  end;

  s := UTF8Encode(StrPas(PWideChar(@FBuffer[idx])));
  ValueGrid.Cells[1, VALUE_ROW_PWIDECHAR] := s;
  ValueGrid.Cells[2, VALUE_ROW_PWIDECHAR] := Format('%d ... %d', [idx, idx + Length(s)]);
end;


procedure TMainForm.ReadCmdLine;
begin
  if ParamCount > 0 then
    LoadFile(ParamStr(1));
end;


procedure TMainForm.ReadFromIni;
var
  ini: TCustomIniFile;
  i: Integer;
begin
  ini := CreateIni;
  try
    ReadFormFromIni(ini, 'MainForm', self);

    TreePanel.Width := ini.ReadInteger('MainForm', 'RecordList_Width', TreePanel.Width);
    for i:=0 to BiffTree.Header.Columns.Count-1 do
      BiffTree.Header.Columns[i].Width := ini.ReadInteger('MainForm',
        Format('RecordList_ColWidth_%d', [i+1]), BiffTree.Header.Columns[i].Width);

    ValueGrid.Height := ini.ReadInteger('MainForm', 'ValueGrid_Height', ValueGrid.Height);
    for i:=0 to ValueGrid.ColCount-1 do
      ValueGrid.ColWidths[i] := ini.ReadInteger('MainForm',
        Format('ValueGrid_ColWidth_%d', [i+1]), ValueGrid.ColWidths[i]);

    for i:=0 to FAnalysisGrid.ColCount-1 do
      FAnalysisGrid.ColWidths[i] := ini.ReadInteger('MainForm',
        Format('AnalysisGrid_ColWidth_%d', [i+1]), FAnalysisGrid.ColWidths[i]);

    AnalysisDetails.Height := ini.ReadInteger('MainForm', 'AnalysisDetails_Height', AnalysisDetails.Height);

    HexEditor.AddressMode := TKHexEditorAddressMode(ini.ReadInteger('HexEditor',
      'AddressMode', ord(HexEditor.AddressMode)));
    CbHexAddress.Checked := HexEditor.AddressMode = eamHex;
    CbHexAddressChange(nil);

    HexEditor.DigitGrouping := ini.ReadInteger('HexEditor',
      'DigitGrouping', HexEditor.DigitGrouping);
    CbHexSingleBytes.Checked := HexEditor.DigitGrouping = 1;
    CbHexSingleBytesChange(nil);

    HexEditor.LineSize := ini.ReadInteger('HexEditor',
      'LineSize', HexEditor.LineSize);
    if HexEditor.LineSize in [10, 16] then
      CbHexEditorLineSize.ItemIndex := 0 else CbHexEditorLineSize.ItemIndex := 1;
    CbHexEditorLineSizeChange(nil);

    PageControl.ActivePageIndex := ini.ReadInteger('MainForm', 'PageIndex', PageControl.ActivePageIndex);
  finally
    ini.Free;
  end;
end;


procedure TMainForm.ReadFromStream(AStream: TStream);
var
  recType: Word;
  recSize: Word;
  p: Cardinal;
  p0: Cardinal;
  s: String;
  i: Integer;
  node, prevnode: PVirtualNode;
  parentnode: PVirtualNode;
  parentdata, data, prevdata: PBiffNodeData;
  w: word;
  crs: TCursor;
begin
  crs := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;
    BiffTree.Clear;
    parentnode := nil;
    FXFIndex := -1;
    FFontIndex := -1;
    FFormatIndex := -1;
    FRowIndex := -1;
    FExternSheetIndex := 0;  // 1-based!
    AStream.Position := 0;
    while AStream.Position < AStream.Size do begin
      p := AStream.Position;
      recType := WordLEToN(AStream.ReadWord);
      recSize := WordLEToN(AStream.ReadWord);
      if (recType = 0) and (recSize = 0) then
        break;
      s := RecTypeName(recType);
      i := pos(':', s);
      // in case of BOF record: create new parent node for this substream
      if (recType = $0009) or (recType = $0209) or (recType = $0409) or (recType = $0809)
      then begin
        // Read info on substream beginning here
        p0 := AStream.Position;
        AStream.Position := AStream.Position + 2;
        w := WordLEToN(AStream.ReadWord);
        AStream.Position := p0;
        // add parent node for this substream
        parentnode := BiffTree.AddChild(nil);
        // add data to parent node
        parentdata := BiffTree.GetNodeData(parentnode);
        BiffTree.ValidateNode(parentnode, False);
        parentdata^.Offset := p;
        parentdata^.RecordName := BOFName(w);
        FRowIndex := -1;
      end;
      // add node to parent node
      node := BIFFTree.AddChild(parentnode);
      data := BiffTree.GetNodeData(node);
      BiffTree.ValidateNode(node, False);
      data^.Offset := p;
      data^.RecordID := recType;
      if i > 0 then begin
        data^.RecordName := copy(s, 1, i-1);
        data^.RecordDescription := copy(s, i+2, Length(s));
      end else begin
        data^.RecordName := s;
        data^.RecordDescription := '';
      end;
      case recType of
        $0008, $0208:  // Row
          begin
            inc(FRowIndex);
            data^.Index := FRowIndex;
          end;
        $0031, $0231:  // Font record
          begin
            inc(FFontIndex);
            if FFontIndex > 3 then data^.Index := FFontIndex + 1
              else data^.Index := FFontIndex;
          end;
        $0043, $00E0:  // XF record
          begin
            inc(FXFIndex);
            data^.Index := FXFIndex;
          end;
        $0017:   // EXTERNSHEET record
          if FFormat < sfExcel8 then begin
            inc(FExternSheetIndex);
            data^.Index := FExternSheetIndex;
          end;
        $001E, $041E:  // Format record
          begin
            inc(FFormatIndex);
            data^.Index := FFormatIndex;
          end;
        $003C:  // CONTINUE reocrd
          begin
            prevnode := BIFFTree.GetPrevious(node);
            prevdata := BiffTree.GetNodeData(prevnode);
            case prevdata^.RecordID of
              $00FC: data^.Tag := BIFFNODE_SST_CONTINUE;    // SST record
              $01B6: data^.Tag := BIFFNODE_TXO_CONTINUE1;   // TX0 record
              $003C: begin                                  // CONTINUE record
                       prevnode := BiffTree.GetPrevious(prevnode);
                       prevdata := BiffTree.GetNodeData(prevnode);
                       if prevdata^.RecordID = $01B6 then   // TX0 record
                         data^.Tag := BIFFNODE_TXO_CONTINUE2;
                     end;
            end;
          end;
        else
          data^.Index := -1;
      end;

      // advance stream pointer
      AStream.Position := AStream.Position + recSize;
    end;

    // expand all parent nodes
    node := BiffTree.GetFirst;
    while node <> nil do begin
      BiffTree.Expanded[node] := true;
      node := BiffTree.GetNextSibling(node);
    end;
    // Select first node
    BiffTree.FocusedNode := BiffTree.GetFirst;
    BiffTree.Selected[BiffTree.FocusedNode] := true;

    UpdateCmds;

  finally
    Screen.Cursor := crs;
  end;
end;


procedure TMainForm.PageControlChange(Sender: TObject);
var
  sel: TKHexEditorSelection;
  i, n: Integer;
  s: String;
begin
  if (BiffTree.FocusedNode = nil) or
     (BiffTree.FocusedNode^.Parent = BiffTree.RootNode)
  then
    exit;

  PopulateAnalysisGrid;
  for i:=1 to FAnalysisGrid.RowCount-1 do begin
    s := FAnalysisGrid.Cells[0, i];
    if s = '' then break;
    n := StrToInt(s);
    if (n >= FCurrOffset) then
    begin
      FAnalysisGrid.Row := IfThen(n = FCurrOffset, i, i-1);
      break;
    end;
  end;

  sel.Index := FCurrOffset;
  sel.Digit := 0;
  HexEditor.SelStart := sel;
  PopulateValueGrid;
  ValueGridClick(nil);
{
  if PageControl.ActivePage = PgAnalysis then
    PopulateAnalysisGrid
  else
  if PageControl.ActivePage = PgValues then
  begin
    sel.Index := FCurrOffset;
    sel.Digit := 0;
    HexEditor.SelStart := sel;
    PopulateValueGrid;
    ValueGridClick(nil);
  end;}
end;


procedure TMainForm.UpdateCaption;
begin
  if FFileName = '' then
    Caption := 'BIFF Explorer - (no file loaded)'
  else
    Caption := Format('BIFF Explorer - "%s [%s]', [
      FFileName,
      GetFileFormatName(FFormat)
    ]);
end;


procedure TMainForm.UpdateCmds;
begin
  AcDumpToFile.Enabled := FFileName <> '';
  AcFind.Enabled := FFileName <> '';
end;


procedure TMainForm.UpdateStatusbar;
begin
  if FCurrOffset > -1 then
    Statusbar.Panels[3].Text := Format('HexViewer offset: %d', [FCurrOffset])
  else
    Statusbar.Panels[3].Text := '';
end;


procedure TMainForm.ValueGridClick(Sender: TObject);
var
  sel: TKHexEditorSelection;
  n: Integer;
begin
  sel := HexEditor.SelStart;

  n := GetValueGridDataSize;
  if n > 0 then begin
    sel.Digit := 0;
    HexEditor.SelStart := sel;
    inc(sel.Index, n-1);
    sel.Digit := 1;
    HexEditor.SelEnd := sel;
  end else
    HexEditor.SelEnd := sel;
end;


procedure TMainForm.ValueGridPrepareCanvas(sender: TObject; aCol,
  aRow: Integer; aState: TGridDrawState);
begin
  if ARow = 0 then ValueGrid.Canvas.Font.Style := [fsBold];
end;


procedure TMainForm.WriteToIni;
var
  ini: TCustomIniFile;
  i: Integer;
begin
  ini := CreateIni;
  try
    WriteFormToIni(ini, 'MainForm', self);

    ini.WriteInteger('MainForm', 'RecordList_Width', TreePanel.Width);
    for i:=0 to BiffTree.Header.Columns.Count-1 do
      ini.WriteInteger('MainForm', Format('RecordList_ColWidth_%d', [i+1]), BiffTree.Header.Columns[i].Width);

    ini.WriteInteger('MainForm', 'ValueGrid_Height', ValueGrid.Height);
    for i:=0 to ValueGrid.ColCount-1 do
      ini.WriteInteger('MainForm', Format('ValueGrid_ColWidth_%d', [i+1]), ValueGrid.ColWidths[i]);

    for i:=0 to FAnalysisGrid.ColCount-1 do
      ini.WriteInteger('MainForm', Format('AnalysisGrid_ColWidth_%d', [i+1]), FAnalysisGrid.ColWidths[i]);

    ini.WriteInteger('MainForm', 'AnalysisDetails_Height', AnalysisDetails.Height);

    ini.WriteInteger('MainForm', 'PageIndex', PageControl.ActivePageIndex);

    ini.WriteInteger('HexEditor', 'AddressMode', ord(HexEditor.AddressMode));
    ini.WriteInteger('HexEditor', 'DigitGrouping', HexEditor.DigitGrouping);
    ini.WriteInteger('HexEditor', 'LineSize', HexEditor.LineSize);

  finally
    ini.Free;
  end;
end;

end.

