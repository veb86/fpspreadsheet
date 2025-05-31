program chart_linestyles;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;
const
  FILE_NAME = 'linestyles_fps';
  nStyles = 8;
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  dir, fn: String;
  i: Integer;
  linestyles: array[1..nStyles] of TsChartLinePatternStyle;
  linestyleNames: array[1..nStyles] of string[20];
  lineWidth: Double;
  fs: TFormatSettings;
begin
  fs := DefaultFormatSettings;
  lineWidth := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  if ParamCount > 0 then
  begin
    if not TryStrToFloat(Paramstr(1), lineWidth) then
    begin
      fs := DefaultFormatSettings;
      if DefaultFormatSettings.DecimalSeparator = '.' then
        fs.DecimalSeparator := ','
      else
        fs.DecimalSeparator := '.';
      TryStrToFloat(ParamStr(1), lineWidth, fs);
    end;
  end;

  fn := Format('%s_%.1gmm', [FILE_NAME, linewidth], fs);

  dir := ExtractFilePath(ParamStr(0)) + 'files/';
  ForceDirectories(dir);

  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('linestyles');
    for i := 0 to 100 do
      sheet.WriteRowHeight(i, 15, suPoints);
//    sheet.WriteDefaultRowHeight(14.5, suMillimeters);

    // Create chart: left/top in cell D2, 100 mm x 160 mm
    ch := sheet.AddChart(100, 160, 3, 1);

    linestyles[1] := clsSolid;                linestyleNames[1] := 'clsSolid';
    linestyles[2] := clsFineDot;              linestyleNames[2] := 'clsFineDot';
    linestyles[3] := clsDot;                  linestyleNames[3] := 'clsDot';
    linestyles[4] := clsDash;                 linestyleNames[4] := 'clsDash';
    linestyles[5] := clsDashDot;              linestyleNames[5] := 'clsDashDot';
    linestyles[6] := clsLongDash;             linestyleNames[6] := 'clsLongDash';
    linestyles[7] := clsLongDashDot;          linestyleNames[7] := 'clsLongDashDot';
    linestyles[8] := clsLongDashDotDot;       linestyleNames[8] := 'clsLongDashDotDot';

    // Enter data
    sheet.WriteNumber(1, 0, 0);
    sheet.WriteNumber(2, 0, 10);
    for i := 1 to nStyles do
    begin
      sheet.WriteText(0, i, linestyleNames[i]);
      sheet.WriteNumber(1, i, nStyles-i + 1);
      sheet.WriteNumber(2, i, nStyles-i + 1);
    end;

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'Chart Linestyles' + LineEnding + Format('Line width = %.1fmm', [linewidth]);
    ch.Title.Font.Style := [fssBold];
    ch.Title.Font.Color := scBlue;
    ch.Legend.Border.Style := clsNoLine;
    ch.Legend.Position := legBottom;
    ch.XAxis.MajorGridLines.Style := clsNoLine;
    ch.XAxis.MinorGridLines.Style := clsNoLine;
    ch.YAxis.MajorGridLines.Style := clsNoLine;
    ch.YAxis.MinorGridLines.Style := clsNoLine;
    ch.XAxis.Max := 10;
    ch.XAxis.Min := 0;
    ch.YAxis.Max := nStyles + 1;
    ch.YAxis.Min := 0;

    for i := 1 to nStyles do
    begin
      ser := TsScatterSeries.Create(ch);
      ser.Line.Style := linestyles[i];
      ser.Line.Width := lineWidth;
      ser.SetXRange(1, 0, 2, 0);
      ser.SetYRange(1, i, 2, i);
      ser.SetTitleAddr(0, i);
    end;

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

