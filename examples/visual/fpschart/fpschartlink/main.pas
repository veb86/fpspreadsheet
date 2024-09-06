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
    Chart1: TChart;
    ComboBox1: TComboBox;
    Label1: TLabel;
    ListChartSource1: TListChartSource;
    Memo1: TMemo;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    sWorkbookSource1: TsWorkbookSource;
    sWorksheetGrid1: TsWorksheetGrid;
    procedure btnBrowseClick(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure ComboBox1CloseUp(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure sWorkbookSource1Error(Sender: TObject; const AMsg: String);
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
  fn := ExpandFileName(Combobox1.Text);
  OpenDialog1.InitialDir := ExtractFilePath(fn);
  OpenDialog1.FileName := '';
  if OpenDialog1.Execute then
  begin
    Combobox1.Text := OpenDialog1.FileName;
    LoadFile(OpenDialog1.FileName);
  end;
end;

procedure TForm1.btnOpenClick(Sender: TObject);
var
  fn: String;
begin
  if FileNameIsAbsolute(Combobox1.Text) then
    fn := Combobox1.Text
  else
    fn := FDir + Combobox1.Text;
  LoadFile(fn);
end;

procedure TForm1.ComboBox1CloseUp(Sender: TObject);
begin
  if ComboBox1.ItemIndex > -1 then
  begin
    Combobox1.Text := FDir + Combobox1.Items[Combobox1.ItemIndex];
    LoadFile(Combobox1.Text);
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
    Combobox1.Items.Assign(L);
  finally
    L.Free;
  end;

  {$IF LCL_FullVersion >= 2020000}
  ComboBox1.TextHint := 'Enter or select file name';
  {$IFEND}
  if ParamCount > 0 then
  begin
    Combobox1.Text := ParamStr(1);
    LoadFile(Combobox1.Text);
  end;
end;

procedure TForm1.sWorkbookSource1Error(Sender: TObject; const AMsg: String);
begin
  Memo1.Lines.Add(AMsg);
end;

procedure TForm1.LoadFile(AFileName: String);
var
  ext: String;
  fn: String;
  i: Integer;
begin
  Memo1.Lines.Clear;

  fn := ExpandFileName(AFileName);
  if not FileExists(fn) then
  begin
    MessageDlg('File "' + fn + '" not found.', mtError, [mbOK], 0);
    exit;
  end;

  ext :=Lowercase(ExtractFileExt(fn));
  if ext = '.ods' then
    sWorkbookSource1.FileFormat := sfOpenDocument
  else
    sWorkbookSource1.Fileformat := sfOOXML;
  sWorkbookSource1.Filename := fn;

  for i := 1 to sWorksheetGrid1.Worksheet.GetLastRowIndex+1 do
    sWorksheetGrid1.AutoRowHeight(1);

  FreeAndNil(sChartLink);

  sChartLink := TsWorkbookChartLink.Create(self);
  sChartLink.Chart := Chart1;
  sChartLink.WorkbookSource := sWorkbookSource1;
  sChartLink.WorkbookChartIndex := 0;
end;

end.

