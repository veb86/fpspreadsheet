unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  TAGraph,
  fpSpreadsheet, fpsTypes, fpsOpenDocument,
  fpSpreadsheetCtrls, fpSpreadsheetGrid,  fpSpreadsheetChart;

type

  { TForm1 }

  TForm1 = class(TForm)
    Chart1: TChart;
    Splitter1: TSplitter;
    sWorkbookSource1: TsWorkbookSource;
    sWorksheetGrid1: TsWorksheetGrid;
    procedure FormCreate(Sender: TObject);
  private
    sChartLink: TsWorkbookChartLink;

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

const
  FILE_NAME = '../../../other/chart/bars.ods';
//  FILE_NAME = '../../../other/chart/area.ods';
//  FILE_NAME = '../../../other/chart/area-sameImg.ods';
//  FILE_NAME = '../../../other/chart/pie.ods';
//  FILE_NAME = '../../../other/chart/scatter.ods';
//  FILE_NAME = '../../../other/chart/regression.ods';
//  FILE_NAME = '../../../other/chart/radar.ods';

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
var
  fn: String;
begin
  fn := ExpandFileName(FILE_NAME);

  sWorkbookSource1.FileFormat := sfOpenDocument;
  if FileExists(fn) then
    sWorkbookSource1.Filename := fn;

  sChartLink := TsWorkbookChartLink.Create(self);
  sChartLink.Chart := Chart1;
  sChartLink.WorkbookSource := sWorkbookSource1;
  sChartLink.WorkbookChartIndex := 0;
end;

end.

