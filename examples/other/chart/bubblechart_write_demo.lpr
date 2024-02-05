program bubblechart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'bubble';
  sDISTANCE = 'Distance from Sun' + LineEnding + '(relative to Earth)';
  sPERIOD = 'Orbital period' + LineEnding + '(relative to Earth)';
  sDIAMETER = 'Diameter (km)';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsBubbleSeries;
begin
  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('bubble_series');

    // Enter data
    sheet.WriteText( 0,0, 'Solar System');
    sheet.WriteFont( 0,0, '', 12, [fssBold], scBlack);

    sheet.WriteText( 2,0, 'Planet' );  sheet.WriteText  ( 2,1, sDistance);  sheet.WriteText  ( 2,2, sPERIOD);  sheet.WriteText  (2, 3, sDIAMETER);
    sheet.WriteText( 3,0, 'Mercury');  sheet.WriteNumber( 3,1, 0.387);      sheet.WriteNumber( 3,2,  0.241);   sheet.WriteNumber( 3,3, 4.879E3);
    sheet.WriteText( 4,0, 'Venus'  );  sheet.WriteNumber( 4,1, 0.723);      sheet.WriteNumber( 4,2,  0.615);   sheet.WriteNumber( 4,3, 1.210E4);
    sheet.WriteText( 5,0, 'Earth'  );  sheet.WriteNumber( 5,1, 1.000);      sheet.WriteNumber( 5,2,  1.000);   sheet.WriteNumber( 5,3, 1.276E4);
    sheet.WriteText( 6,0, 'Mars'   );  sheet.WriteNumber( 6,1, 1.524);      sheet.WriteNumber( 6,2,  1.881);   sheet.WriteNumber( 6,3, 6.792E3);
    sheet.WriteText( 7,0, 'Jupiter');  sheet.WriteNumber( 7,1, 5.204);      sheet.WriteNumber( 7,2, 11.862);   sheet.WriteNumber( 7,3, 1.430E5);
    sheet.WriteText( 8,0, 'Saturn' );  sheet.WriteNumber( 8,1, 9.582);      sheet.WriteNumber( 8,2, 29.445);   sheet.WriteNumber( 8,3, 1.205E5);
    sheet.WriteText( 9,0, 'Uranus' );  sheet.WriteNumber( 9,1,19.201);      sheet.WriteNumber( 9,2, 84.011);   sheet.WriteNumber( 9,3, 5.112E4);
    sheet.WriteText(10,0, 'Neptune');  sheet.WriteNumber(10,1,30.047);      sheet.WriteNumber(10,2,164.79);    sheet.WriteNumber(10,3, 4.953E4);

    sheet.WriteText(12,0, 'Source: wikipedia');
    sheet.WriteFont(12,0, '', 8, [], scBlack);

    sheet.WriteColWidth(1, 40.0, suMillimeters);
    sheet.WriteColWidth(2, 40.0, suMillimeters);

    // Create chart: left/top in cell D4, 150 mm x 150 mm
    ch := book.AddChart(sheet, 2, 4, 150, 150);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'Solar System';
    ch.Title.Font.Style := [fssBold];
    ch.Title.Font.Color := scBlue;
    ch.Legend.Visible := false;
    ch.XAxis.Title.Caption := 'Distance from Sun (relative to Earth)';
    ch.XAxis.MinorGridLines.Style := clsNoLine;
    ch.XAxis.Logarithmic := true;
    ch.XAxis.Min := 0.1;
    ch.XAxis.Max := 100;
    ch.YAxis.Title.Caption := 'Orbital period (relative to Earth)';
    ch.YAxis.AxisLine.Color := scSilver;
    ch.YAxis.MinorGridLines.Style := clsNoLine;
    ch.YAxis.Logarithmic := true;
    ch.YAxis.Min := 0.1;
    ch.YAxis.Max := 1000;

    // Add data as bubble series
    ser := TsBubbleSeries.Create(ch);
    ser.SetLabelRange(3, 0, 10, 0);
    ser.SetXRange(3, 1, 10, 1);
    ser.SetYRange(3, 2, 10, 2);
    ser.SetBubbleRange(3, 3, 10, 3);
    ser.Line.Style := clsSolid; //NoLine;
    ser.Line.Color := scSilver;
    ser.Fill.Color := scYellow;
    ser.Fill.Transparency := 0.25;
    ser.DataLabels := [cdlCategory];
    ser.DataPointStyles.AddSolidFill(2, $c47244);
    ser.DataPointStyles.AddSolidFill(3, scRed);

    book.WriteToFile(FILE_NAME + '.xlsx', true);
    WriteLn('Data saved with chart in ', FILE_NAME + '.xlsx');

    book.WriteToFile(FILE_NAME + '.ods', true);
    WriteLn('Data saved with chart in ', FILE_NAME + '.ods');
  finally
    book.Free;
  end;
end.

