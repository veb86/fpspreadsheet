unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  LCLVersion, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls, FileUtil, LazFileUtils,
  TAGraph, TASources,
  fpSpreadsheet, fpsTypes, fpsOpenDocument, xlsxOOXML,
  fpSpreadsheetCtrls, fpSpreadsheetGrid, fpSpreadsheetChart;

type

  { TForm1 }

  TForm1 = class(TForm)
    btnBrowse: TButton;
    btnOpen: TButton;
    Chart: TChart;
    cbFileNames: TComboBox;
    lblFileNames: TLabel;
    ListChartSource: TListChartSource;
    Memo: TMemo;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    sWorkbookSource: TsWorkbookSource;
    sWorksheetGrid: TsWorksheetGrid;
    procedure btnBrowseClick(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure cbFileNamesSelect(Sender:TObject);
    procedure FormCreate(Sender: TObject);
    procedure sWorkbookSourceError(Sender: TObject; const AMsg: String);
  private
    FDir: String;
    sChartLink: TsWorkbookChartLink;
    procedure LoadFile(AFileName: String);

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

uses
  TypInfo,
  TAChartUtils, TAChartAxis, TAChartAxisUtils, TACustomSeries, TATransformations;


{ TForm1 }

procedure TForm1.btnBrowseClick(Sender: TObject);
var
  fn: String;
begin
  fn := ExpandFileName(cbFileNames.Text);
  OpenDialog.InitialDir := ExtractFilePath(fn);
  OpenDialog.FileName := '';
  if OpenDialog.Execute then
  begin
    cbFileNames.Text := OpenDialog.FileName;
    LoadFile(OpenDialog.FileName);
  end;
end;

procedure TForm1.btnOpenClick(Sender: TObject);
var
  fn: String;
begin
  if FileNameIsAbsolute(cbFileNames.Text) then
    fn := cbFileNames.Text
  else
    fn := FDir + cbFileNames.Text;
  LoadFile(fn);
end;

procedure TForm1.cbFileNamesSelect(Sender:TObject);
var
  fn: String;
begin
  if cbFileNames.ItemIndex > -1 then
  begin
    fn := FDir + cbFileNames. Items[cbFileNames.ItemIndex];
    LoadFile(fn);
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  L: TStrings;
  i: Integer;
begin
  FDir := ExpandFileName(Application.Location + '../../../other/chart/files/');
  L := TStringList.Create;
  try
    FindAllFiles(L, FDir, '*.xlsx;*.ods', false);
    for i := 0 to L.Count-1 do
      L[i] := ExtractFileName(L[i]);
    cbFileNames.Items.Assign(L);
  finally
    L.Free;
  end;

  {$IF LCL_FullVersion >= 2020000}
  cbFileNames.TextHint := 'Enter or select file name';
  {$IFEND}
  if ParamCount > 0 then
  begin
    cbFileNames.Text := ParamStr(1);
    LoadFile(cbFileNames.Text);
  end;
end;

procedure TForm1.sWorkbookSourceError(Sender: TObject; const AMsg: String);
begin
  Memo.Lines.Add(AMsg);
end;

procedure TForm1.LoadFile(AFileName: String);
var
  ext: String;
  fn: String;
  i: Integer;
begin
  Memo.Lines.Clear;

  fn := ExpandFileName(AFileName);
  if not FileExists(fn) then
  begin
    MessageDlg('File "' + fn + '" not found.', mtError, [mbOK], 0);
    exit;
  end;

  ext :=Lowercase(ExtractFileExt(fn));
  if ext = '.ods' then
    sWorkbookSource.FileFormat := sfOpenDocument
  else
    sWorkbookSource.Fileformat := sfOOXML;
  sWorkbookSource.Filename := fn;

  for i := 1 to sWorksheetGrid.Worksheet.GetLastRowIndex+1 do
    sWorksheetGrid.AutoRowHeight(1);

  FreeAndNil(sChartLink);

  sChartLink := TsWorkbookChartLink.Create(self);
  sChartLink.Chart := Chart;
  sChartLink.WorkbookSource := sWorkbookSource;
  sChartLink.WorkbookChartIndex := 0;
end;

end.

