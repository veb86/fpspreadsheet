program radarchart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'radar';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsRadarSeries;
  fn, dir: String;
begin
  dir := ExtractFilePath(ParamStr(0)) + 'files/';
  ForceDirectories(dir);
  fn := dir + FILE_NAME;

  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('radar_series');

    // Enter data
    sheet.WriteText( 0, 0, 'School Grades');
    sheet.WriteFont( 0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText( 2, 0, '');          sheet.WriteText  ( 2, 1, 'Student 1'); sheet.WriteText  ( 2, 2, 'Student 2');
    sheet.WriteText( 3, 0, 'Biology');   sheet.WriteNumber( 3, 1, 12);          sheet.WriteNumber( 3, 2, 15);
    sheet.WriteText( 4, 0, 'History');   sheet.WriteNumber( 4, 1, 11);          sheet.WriteNumber( 4, 2, 13);
    sheet.WriteText( 5, 0, 'French');    sheet.WriteNumber( 5, 1, 16);          sheet.WriteNumber( 5, 2, 11);
    sheet.WriteText( 6, 0, 'English');   sheet.WriteNumber( 6, 1, 18);          sheet.WriteNumber( 6, 2, 11);
    sheet.WriteText( 7, 0, 'Sports');    sheet.WriteNumber( 7, 1, 16);          sheet.WriteNumber( 7, 2,  7);
    sheet.WriteText( 8, 0, 'Maths');     sheet.WriteNumber( 8, 1, 10);          sheet.WriteNumber( 8, 2, 17);
    sheet.WriteText( 9, 0, 'Physics');   sheet.WriteNumber( 9, 1, 12);          sheet.WriteNumber( 9, 2, 19);
    sheet.WriteText(10, 0, 'Computer');  sheet.WriteNumber(10, 1, 16);          sheet.WriteNumber(10, 2, 18);

    // Create chart: left/top in cell D4, 160 mm x 100 mm
    ch := book.AddChart(sheet, 2, 3, 120, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'School Grades';
    ch.Title.Font.Style := [fssBold];
    ch.Legend.Border.Style := clsNoLine;
    ch.XAxis.Title.Caption := '';
    ch.YAxis.Title.Caption := '';
    ch.YAxis.AxisLine.Color := scSilver;
    ch.YAxis.MajorTicks := [];

    // Add 1st radar series ("Student 1")
    ser := TsRadarSeries.Create(ch);
    ser.SetTitleAddr(2, 1);
    ser.SetLabelRange(3, 0, 10, 0);
    ser.SetYRange(3, 1, 10, 1);
    ser.Line.Color := scDarkRed;
    //ser.Fill.Style := cfsNoFill;    // --> non-filled radar chrt
    ser.ShowSymbols := true;
    ser.Symbol := cssDiamond;
    ser.SymbolFill.Style := cfsSolid;
    ser.SymbolFill.Color := scYellow;
    // in ods the symbol color is always equal to the line color
    ser.SymbolWidth := 12; //3;
    ser.SymbolHeight := 12; // 3;

    // Add 2nd radar series ("Student 2")
    ser := TsFilledRadarSeries.Create(ch);
    ser.SetTitleAddr(2, 2);
    ser.SetLabelRange(3, 0, 10, 0);
    ser.SetYRange(3, 2, 10, 2);
    ser.Line.Color := scDarkBlue;
    ser.Fill.Color := $FFCC99;
    ser.Fill.Transparency := 0.35;

    book.WriteToFile(fn + '.xlsx', true);
    WriteLn('Data saved with chart in ', fn + '.xlsx');

    book.WriteToFile(fn + '.ods', true);
    WriteLn('Data saved with chart in ', fn + '.ods');
  finally
    book.Free;
  end;
end.

