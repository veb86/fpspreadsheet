program barchart_stacked_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

procedure WriteHelp;
begin
  WriteLn('SYNTAX: barchart_stacked_write_demo [rotated]');
  WriteLn('  rotated ........... hoizontal bars, otherwise: vertical bar');
  WriteLn('  percentage ........ stacked percentage, otherwise: normal stacking');
  Halt;
end;

const
  FILE_NAME = 'bars-stacked';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  fn: String;
  i: Integer;
  rotated: Boolean = false;
  stackedPercentage: Boolean = false;
begin
  if (ParamCount = 1) and ((ParamStr(1) = '--help') or (ParamStr(1) = '-h')) then
    WriteHelp;

  fn := FILE_NAME;
  for i := 1 to ParamCount do
    case lowercase(ParamStr(i)) of
      'rotated':
        begin
          rotated := true;
          fn := fn + '-rotated';
        end;
      'percentage':
        begin
          stackedPercentage := true;
          fn := fn + '-percentage';
        end;
    end;

  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('bar_series_stacked');

    // Enter data
    sheet.WriteText( 0, 0, 'Stacked bar series demo');
    sheet.WriteFont( 0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText( 2, 0, 'Quarters'); sheet.WriteText  ( 2, 1, 'Product A'); sheet.WriteText  ( 2, 2, 'Product B');
    sheet.WriteText( 3, 0, 'Q1/2022');  sheet.WriteNumber( 3, 1, 125);         sheet.WriteNumber( 3, 2, 207);
    sheet.WriteText( 4, 0, 'Q2/2022');  sheet.WriteNumber( 4, 1, 176);         sheet.WriteNumber( 4, 2, 199);
    sheet.WriteText( 5, 0, 'Q3/2022');  sheet.WriteNumber( 5, 1, 264);         sheet.WriteNumber( 5, 2, 194);
    sheet.WriteText( 6, 0, 'Q4/2022');  sheet.WriteNumber( 6, 1, 311);         sheet.WriteNumber( 6, 2, 183);

    // Create chart: left/top in cell D4, 160 mm x 100 mm
    ch := book.AddChart(sheet, 2, 3, 120, 100);

    // Chart properties
    if stackedPercentage then
      ch.StackMode :=csmStackedPercentage
    else
      ch.StackMode := csmStacked;
    if rotated then
      ch.RotatedAxes := rotated;
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'Product Sales';
    ch.Title.Font.Style := [fssBold];
    ch.Legend.Border.Style := clsNoLine;
    ch.XAxis.Title.Caption := '';
    ch.YAxis.Title.Caption := 'Quarter';
    ch.YAxis.AxisLine.Color := scSilver;
    ch.YAxis.MajorTicks := [];

    // Add 1st bar series ("Product A")
    ser := TsBarSeries.Create(ch);
    ser.SetTitleAddr(2, 1);
    ser.SetLabelRange(3, 0, 6, 0);
    ser.SetYRange(3, 1, 6, 1);
    ser.Fill.Color := $3810F3;
    ser.Line.Style := clsNoLine;

    // Add 2nd bar series ("Product B")
    ser := TsBarSeries.Create(ch);
    ser.SetTitleAddr(2, 2);
    ser.SetLabelRange(3, 0, 6, 0);
    ser.SetYRange(3, 2, 6, 2);
    ser.Fill.Color := $4200A8;
    ser.Line.Style := clsNoLine;

    book.WriteToFile(fn + '.xlsx', true);
    WriteLn('Data saved with chart in ', fn + '.xlsx');

    book.WriteToFile(fn + '.ods', true);
    WriteLn('Data saved with chart in ', fn + '.ods');
  finally
    book.Free;
  end;
end.

