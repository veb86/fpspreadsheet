program piechart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'pie';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  fill: TsChartFill;
  line: TsChartLine;
begin
  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('pie_series');

    // Enter data
    sheet.WriteText(0, 0, 'World population');
    sheet.WriteFont(0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText(1, 0, 'https://en.wikipedia.org/wiki/World_population');
    sheet.WriteHyperlink(1, 0, 'https://en.wikipedia.org/wiki/World_population');
    sheet.WriteText(3, 0, 'Continent');  sheet.WriteText  (3, 1, 'Population (millions)');
    sheet.WriteFontStyle(3, 0, [fssBold]); sheet.WriteFontStyle(3, 1, [fssBold]);
    sheet.WriteText(4, 0, 'Asia');       sheet.WriteNumber(4, 1, 4641);      // sheet.WriteChartColor(4, 2, scYellow);
    sheet.WriteText(5, 0, 'Africa');     sheet.WriteNumber(5, 1, 1340);      // sheet.WriteChartColor(5, 2, scBrown);
    sheet.WriteText(6, 0, 'America');    sheet.WriteNumber(6, 1, 653 + 368); // sheet.WriteChartColor(6, 2, scRed);
    sheet.WriteText(7, 0, 'Europe');     sheet.WriteNumber(7, 1, 747);       // sheet.WriteChartColor(7, 2, scSilver);
    sheet.WriteText(8, 0, 'Oceania');    sheet.WriteNumber(8, 1, 42);        // sheet.WriteChartColor(8, 2, $FF8080);

    // Create chart: left/top in cell D4, 150 mm x 150 mm
    ch := book.AddChart(sheet, 2, 3, 150, 150);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'World Population';
    ch.Title.Font.Style := [fssBold];
    ch.SubTitle.Caption := '(in millions)';
    ch.SubTitle.Font.Size := 10;
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
    ser.LabelFormat := '#,##0';

    // Individual sector colors, with white border
    // Must be complete, otherwise will be ignored by Calc and replaced by default colors
    line := TsChartline.CreateSolid(scWhite, 0.8);
    fill := TsChartFill.CreateHatchFill(ch.Hatches.AddLineHatch('ltHorz', chsSingle, $00C0FF, 1, 0.1, 0), scWhite);
    ser.DataPointStyles.AddSolidFill($C47244, line);
    ser.DataPointStyles.AddSolidFill($317DED, line);
    ser.DataPointStyles.AddSolidFill($A5A5A5, line);
    ser.DataPointStyles.AddFillAndLine(fill, line);
//    ser.DataPointStyles.AddSolidFill($00C0FF, line);
    line.Color := scWhite;
    ser.DataPointStyles.AddSolidFill($D69B5B, line);
    line.Free;

    //ser.SetFillColorRange(4, 2, 8, 2);

    book.WriteToFile(FILE_NAME+'.xlsx', true);
    WriteLn('Data saved with chart in ', FILE_NAME+'.xlsx');

    book.WriteToFile(FILE_NAME + '.ods', true);
    WriteLn('Data saved with chart in ', FILE_NAME+'.ods');
  finally
    book.Free;
  end;
end.

