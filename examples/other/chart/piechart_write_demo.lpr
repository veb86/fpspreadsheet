program piechart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;
var
  b: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
begin
  b := TsWorkbook.Create;
  try
    // worksheet
    sheet := b.AddWorksheet('pie_series');

    // Enter data
    sheet.WriteText(0, 0, 'World population');
    sheet.WriteFont(0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText(1, 0, 'https://en.wikipedia.org/wiki/World_population');
    sheet.WriteHyperlink(1, 0, 'https://en.wikipedia.org/wiki/World_population');
    sheet.WriteText(3, 0, 'Continent');  sheet.WriteText  (3, 1, 'Population (millions)');
    sheet.WriteText(4, 0, 'Asia');       sheet.WriteNumber(4, 1, 4641);      // sheet.WriteChartColor(4, 2, scYellow);
    sheet.WriteText(5, 0, 'Africa');     sheet.WriteNumber(5, 1, 1340);      // sheet.WriteChartColor(5, 2, scBrown);
    sheet.WriteText(6, 0, 'America');    sheet.WriteNumber(6, 1, 653 + 368); // sheet.WriteChartColor(6, 2, scRed);
    sheet.WriteText(7, 0, 'Europe');     sheet.WriteNumber(7, 1, 747);       // sheet.WriteChartColor(7, 2, scSilver);
    sheet.WriteText(8, 0, 'Oceania');    sheet.WriteNumber(8, 1, 42);        // sheet.WriteChartColor(8, 2, $FF8080);

    // Create chart: left/top in cell D4, 120 mm x 100 mm
    ch := b.AddChart(sheet, 2, 3, 120, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'World Population';
    ch.Title.Font.Style := [fssBold];
    ch.Legend.Border.Style := clsNoLine;

    // Add pie series
    ser := TsPieSeries.Create(ch);       // Select one of these...
    //ser := TsRingSeries.Create(ch);

    // Series properties
    ser.SetTitleAddr(0, 0);
    ser.SetLabelRange(4, 0, 8, 0);
    ser.SetYRange(4, 1, 8, 1);
    ser.DataLabels := [cdlCategory, cdlValue];
    ser.LabelSeparator := '\n'; // this is the symbol for a line-break
    ser.LabelPosition := lpOutside;
    ser.Line.Color := scWhite;
    //ser.SetFillColorRange(4, 2, 8, 2);

    b.WriteToFile('world-population.xlsx', true);   // Excel fails to open the file
    b.WriteToFile('world-population.ods', true);
  finally
    b.Free;
  end;
end.
