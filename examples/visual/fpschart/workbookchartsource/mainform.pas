unit mainform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, TAGraph, TASources, TASeries, TAStyles, Forms,
  Controls, Graphics, Dialogs, ExtCtrls, StdCtrls, fpspreadsheetctrls,
  fpspreadsheetgrid, fpspreadsheetchart, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    Bevel1: TBevel;
    BtnDeleteSheet: TButton;
    BtnRenameSheet: TButton;
    Button2: TButton;
    Chart1: TChart;
    Chart1AreaSeries1: TLineSeries;
    Chart1LineSeries1: TLineSeries;
    Chart2: TChart;
    Chart2BarSeries1: TBarSeries;
    Chart3: TChart;
    Chart3PieSeries1: TPieSeries;
    ChartStyles1: TChartStyles;
    Panel1: TPanel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    sWorkbookChartSource1: TsWorkbookChartSource;
    sWorkbookChartSource2: TsWorkbookChartSource;
    sWorkbookChartSource3: TsWorkbookChartSource;
    sWorkbookSource1: TsWorkbookSource;
    sWorkbookTabControl1: TsWorkbookTabControl;
    sWorksheetGrid1: TsWorksheetGrid;
    procedure BtnDeleteSheetClick(Sender: TObject);
    procedure BtnRenameSheetClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
begin
  sWorkbookSource1.Filename := 'test-data.xlsx';

  sWorkbookChartSource1.XRange := 'Sheet1!A2:A32';
  sWorkbookChartSource1.YRange := '(Sheet1!B2:B32) (Sheet1!C2:C32) (Sheet1!D2:D17;Sheet1!E18:E32)';

  sWorkbookChartSource2.XRange := 'Sheet2!A2:A27';
  sWorkbookChartSource2.YRange := 'Sheet2!B2:B27';

  sWorkbookChartSource3.XRange := 'Sheet3!C2:C5';
  sWorkbookChartSource3.YRange := 'Sheet3!B2:B5';
  sWorkbookChartSource3.LabelRange := 'Sheet3!A2:A5';
end;

procedure TForm1.BtnDeleteSheetClick(Sender: TObject);
begin
  if sWorkbookSource1.Workbook.GetWorksheetCount = 1 then
    MessageDlg('There must be a least 1 worksheet.', mtError, [mbOK], 0)
  else
  if MessageDlg('Do you really want to delete this worksheet?', mtConfirmation,
    [mbYes, mbNo], 0) = mrYes
  then
    sWorkbookSource1.Workbook.RemoveWorksheet(sWorkbookSource1.Worksheet);
end;

procedure TForm1.BtnRenameSheetClick(Sender: TObject);
var
  s: String;
begin
  s := sWorkbookSource1.Worksheet.Name;
  if InputQuery('Edit worksheet name', 'New name', s) then
  begin
    if sWorkbookSource1.Workbook.ValidWorksheetName(s) then
      sWorkbookSource1.Worksheet.Name := s
    else
      MessageDlg('Invalid worksheet name.', mtError, [mbOK], 0);
  end;
end;


end.

