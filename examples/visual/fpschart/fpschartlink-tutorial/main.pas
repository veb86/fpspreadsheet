unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, ExtCtrls, SysUtils,
  Forms, Controls, Graphics, Dialogs,
  TAGraph,
  fpspreadsheet, fpstypes, fpsutils, fpschart, xlsxooxml, fpsOpenDocument,
  fpspreadsheetchart, fpspreadsheetctrls, fpspreadsheetgrid;

type
  TForm1 = class(TForm)
    Chart1:TChart;
    Splitter1: TSplitter;
    sWorkbookChartLink1:TsWorkbookChartLink;
    sWorkbookSource1:TsWorkbookSource;
    sWorkbookTabControl1: TsWorkbookTabControl;
    sWorksheetGrid1:TsWorksheetGrid;
    procedure FormCreate(Sender:TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

procedure TForm1.FormCreate(Sender:TObject);
var
  wbook: TsWorkbook;
  wsheet: TsWorksheet;
  wChart: TsChart;
  ser1, ser2: TsChartSeries;
//  ser: TsScatterSeries;
begin
  // Create workbook and worksheet
  wbook := TsWorkbook.Create;
  wsheet := wbook.AddWorksheet('Sales');

  // Add data for plotting
  wsheet.WriteText(0, 0, 'Sales Results');
  wsheet.WriteFontSize(0, 0, 12);
  wsheet.WriteText(1, 0, '(in millions of Euros)');

  wsheet.WriteText(3, 1, 'Product A');
  wsheet.WriteText(3, 2, 'Product B');

  wsheet.WriteText(4, 0, 'London');  wsheet.WriteNumber(4, 1, 1.6);    wsheet.WriteNumber(4, 2, 2.3);
  wsheet.WriteText(5, 0, 'Paris');   wsheet.WriteNumber(5, 1, 1.2);    wsheet.WriteNumber(5, 2, 1.0);
  wsheet.WriteText(6, 0, 'Rome');    wsheet.WriteNumber(6, 1, 1.3);    wsheet.WriteNumber(6, 2, 0.5);

  // Add the chart
  wChart := wbook.AddChart(wsheet, 150, 90, 0, 3, 10);

  // Add first series and set its properties
  ser1 := TsBarSeries.Create(wChart);
  ser1.SetLabelRange(4, 0, 6, 0);   // A5:A7
  ser1.SetYRange(4, 1, 6, 1);       // B5:B7
  ser1.SetTitleAddr(3, 1);          // B4
  ser1.Fill.Color := ChartColor($FFAA88);
  ser1.Line.Color := ChartColor(scBlack);

  // Add second series and set its properties
  ser2 := TsBarSeries.Create(wChart);
  ser2.SetLabelRange(4, 0, 6, 0);   // A5:A7
  ser2.SetYRange(4, 2, 6, 2);       // C5:C7
  ser2.SetTitleAddr(3, 2);          // C4

  ser2.Fill.Pattern := wChart.FillPatterns.AddPattern('white-on-blue diag-up', fpsDiagUpNarrow, ChartColor(scWhite), ChartColor(scBlue));
  ser2.Fill.Style := cfsPattern;
  ser2.Line.Color := ChartColor(scBlack);

  { Alternatively fill by a bitmap:
  ser2.Fill.Image := Chart1.Images.AddImage('FillImg', wbook.AddEmbeddedObj('bgimage.jpg'));
  ser2.Fill.Style := cfsImage;
  ser2.Line.Color := ChartColor(scBlack);
   }

  // No chart border
  wChart.Border.Style := clsNoLine;

  // Chart background gradient
  wChart.Background.Style := cfsGradient;
  wChart.Background.Gradient := wChart.Gradients.AddLinearGradient('gBkGr', ChartColor(scWhite), ChartColor($FFDDAA), 90.0);

  // Show the legend
  wChart.Legend.Visible := true;
  wChart.Legend.Position := lpBottom;

  // Chart title
  wChart.Title.Caption := 'Sales Report';
  wChart.Title.Font.Size := 16;
  wChart.Title.Font.Style := [fssBold];

  // y axis title
  wChart.yAxis.Title.Caption := 'Sales (in millions of EUR)';

  // Make the workbook source load this workbook
  sWorkbookSource1.LoadFromWorkbook(wbook);

  // Save to xlsx and ods files.
  wbook.WriteToFile('test-chart.xlsx', true);
  wbook.WriteToFile('test-chart.ods', true);
end;

end.

