program write_chart_demo;

uses
  SysUtils, fpspreadsheet, fpstypes, fpschart, xlsxooxml, fpsopendocument;
var
  b: TsWorkbook;
  sh1, sh2, sh3: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  i: Integer;
  bg: TsChartFill;
  frm: TsChartLine;
begin
  b := TsWorkbook.Create;
  try
    // 1st sheet
    sh1 := b.AddWorksheet('test1');
    sh1.WriteText(0, 1, 'sin(x)');
    for i := 1 to 7 do
    begin
      sh1.WriteNumber(i, 0, i-1);
      sh1.WriteNumber(i, 1, sin(i-1));
    end;

    ch := b.AddChart(sh1, 4, 4, 125, 95);
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(0, 1);
    ser.SetLabelRange(1, 0, 7, 0);
    ser.SetYRange(1, 1, 7, 1);

    bg.FgColor := scYellow;
    bg.Style := fsSolidFill;
    ch.Background := bg;

    frm.color := scRed;
    frm.Style := clsSolid;
    ch.Border := frm;

    ch.Title.Caption := 'HALLO';
    ch.Title.Visible := true;
    ch.SubTitle.Caption := 'hallo';
    ch.SubTitle.Visible := true;
    ch.YAxis.ShowMajorGridLines := true;
    ch.YAxis.ShowMinorGridLines := true;

    // 2nd sheet
    sh2 := b.AddWorksheet('test2');

    // 3rd sheet
    sh3 := b.AddWorksheet('test3');
    sh3.WriteText(0, 1, 'cos(x)');
    sh3.WriteText(0, 2, 'sin(x)');
    for i := 1 to 7 do
    begin
      sh3.WriteNumber(i, 0, i-1);
      sh3.WriteNumber(i, 1, cos(i-1), nfFixed, 2);
      sh3.WriteNumber(i, 2, sin(i-1), nfFixed, 2);
    end;

    ch := b.AddChart(sh3, 1, 3, 125, 95);
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(0, 1);
    ser.SetLabelRange(1, 0, 7, 0);
    ser.SetYRange(1, 1, 7, 1);
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(0, 2);
    ser.SetLabelRange(1, 0, 7, 0);
    ser.SetYRange(1, 2, 7, 2);
    ch.Title.Caption := 'HALLO';
    ch.Title.Visible := true;
    ch.SubTitle.Caption := 'hallo';
    ch.Subtitle.Visible := true;
    ch.XAxis.ShowMajorGridLines := true;
    ch.XAxis.ShowMinorGridLines := true;

    b.WriteToFile('test.xlsx', true);   // Excel fails to open the file
    b.WriteToFile('test.ods', true);
  finally
    b.Free;
  end;
end.

