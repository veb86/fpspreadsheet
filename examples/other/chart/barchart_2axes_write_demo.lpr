program barchart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'bars_2axes';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
begin
  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('bar_series');

    // Enter data
    sheet.WriteText( 0, 0, 'Test Results');
    sheet.WriteFont( 0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText( 2, 0, '');        sheet.WriteText  ( 2, 1, 'Count'); sheet.WriteText  ( 2, 2, 'Volume');
    sheet.WriteText( 3, 0, 'Case 1');  sheet.WriteNumber( 3, 1, 12);      sheet.WriteNumber( 3, 2,  501);
    sheet.WriteText( 4, 0, 'Case 2');  sheet.WriteNumber( 4, 1, 24);      sheet.WriteNumber( 4, 2, 1054);
    sheet.WriteText( 5, 0, 'Case 3');  sheet.WriteNumber( 5, 1, 21);      sheet.WriteNumber( 5, 2, 4432);
    sheet.WriteText( 6, 0, 'Case 4');  sheet.WriteNumber( 6, 1, 19);      sheet.WriteNumber( 6, 2, 6982);
    sheet.WriteText( 7, 0, 'Case 5');  sheet.WriteNumber( 7, 1,  9);      sheet.WriteNumber( 7, 2,  304);
    sheet.WriteText( 8, 0, 'Case 6');  sheet.WriteNumber( 8, 1,  5);      sheet.WriteNumber( 8, 2, 1285);

    // Create chart: left/top in cell D4, 160 mm x 100 mm
    ch := book.AddChart(sheet, 2, 3, 120, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'Test Results';
    ch.Title.Font.Style := [fssBold];
    ch.Legend.Border.Style := clsNoLine;
    ch.XAxis.Title.Caption := '';
    ch.YAxis.Title.Caption := 'Count';
    ch.YAxis.Title.Font.Color := $0075ea; //597bff;
    ch.YAxis.AxisLine.Color := $0075ea; //597bff;
    ch.YAxis.LabelFont.Color := $0075ea; //597bff;
    ch.YAxis.MajorTicks := [];
    ch.Y2Axis.Title.Caption := 'Volume';
    ch.Y2Axis.Title.Font.Color := $b08359;
    ch.Y2Axis.AxisLine.Color := $b08359;
    ch.Y2Axis.LabelFont.Color := $b08359;

    // Add 1st bar series ("Count")
    ser := TsBarSeries.Create(ch);
    ser.YAxis := alPrimary;
    ser.SetTitleAddr(2, 1);
    ser.SetLabelRange(3, 0, 8, 0);
    ser.SetYRange(3, 1, 8, 1);
    ser.Fill.Style := cfsSolid;
    ser.Fill.Color := $0075ea; //597bff;
    ser.Fill.Color := scRed;
    ser.Line.Style := clsNoLine;

    // Add 2nd bar series ("Volume")
    ser := TsBarSeries.Create(ch);
    ser.YAxis := alSecondary;
    ser.SetTitleAddr(2, 2);
    ser.SetLabelRange(3, 0, 8, 0);
    ser.SetYRange(3, 2, 8, 2);
    ser.Fill.Style := cfsSolid;
    ser.Fill.Color := $b08359;
    ser.Line.Style := clsNoLine;

    {
    book.WriteToFile(FILE_NAME + '.xlsx', true);   // Excel fails to open the file
    WriteLn('Data saved with chart in ', FILENAME, '.xlsx');
    }

    book.WriteToFile(FILE_NAME + '.ods', true);
    WriteLn('Data saved with chart in ', FILE_NAME, '.ods');
  finally
    book.Free;
  end;
end.

