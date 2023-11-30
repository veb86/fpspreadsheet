unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  TAGraph,
  fpSpreadsheet, fpsTypes, fpsOpenDocument,
  fpSpreadsheetCtrls, fpSpreadsheetGrid,  fpSpreadsheetChart;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Chart1: TChart;
    ComboBox1: TComboBox;
    Label1: TLabel;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    Splitter1: TSplitter;
    sWorkbookSource1: TsWorkbookSource;
    sWorksheetGrid1: TsWorksheetGrid;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ComboBox1CloseUp(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    sChartLink: TsWorkbookChartLink;
    procedure LoadFile(AFileName: String);

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

const
//  FILE_NAME = '../../../other/chart/bars.ods';
  FILE_NAME = '../../../other/chart/bubble.ods';
//  FILE_NAME = '../../../other/chart/area.ods';
//  FILE_NAME = '../../../other/chart/area-sameImg.ods';
//  FILE_NAME = '../../../other/chart/pie.ods';
//  FILE_NAME = '../../../other/chart/scatter.ods';
//  FILE_NAME = '../../../other/chart/regression.ods';
//  FILE_NAME = '../../../other/chart/radar.ods';

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
begin
  OpenDialog1.InitialDir := ExtractFilePath(Combobox1.Text);
  OpenDialog1.FileName := '';
  if OpenDialog1.Execute then
  begin
    Combobox1.Text := OpenDialog1.FileName;
    LoadFile(OpenDialog1.FileName);
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  LoadFile(Combobox1.Text);
end;

procedure TForm1.ComboBox1CloseUp(Sender: TObject);
begin
  Combobox1.Text := Combobox1.Items[Combobox1.ItemIndex];
  LoadFile(Combobox1.Text);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  if ParamCount > 0 then
  begin
    Combobox1.Text := ParamStr(1);
    LoadFile(Combobox1.Text);
  end;
end;

procedure TForm1.LoadFile(AFileName: String);
var
  fn: String;
  i: Integer;
begin
  fn := ExpandFileName(AFileName);
  if not FileExists(fn) then
  begin
    MessageDlg('File "' + fn + '" not found.', mtError, [mbOK], 0);
    exit;
  end;

  sWorkbookSource1.FileFormat := sfOpenDocument;
  if FileExists(fn) then
    sWorkbookSource1.Filename := fn;

  for i := 1 to sWorksheetGrid1.Worksheet.GetLastRowIndex+1 do
    sWorksheetGrid1.AutoRowHeight(1);

  sChartLink := TsWorkbookChartLink.Create(self);
  sChartLink.Chart := Chart1;
  sChartLink.WorkbookSource := sWorkbookSource1;
  sChartLink.WorkbookChartIndex := 0;
end;
end.

