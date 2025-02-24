program areachart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

procedure WriteHelp;
begin
  WriteLn('SYNTAX: areachart_write_demo [rotated] [side-by-side|stacked|percent-stacked] ');
  WriteLn('  (no argument) ..... normal orientation (x horizontal)');
  WriteLn('  rotated ........... axes rotated (x vertical)');
  WriteLn('  normal ............ back-to-bottom areas (default)');
  WriteLn('  stacked ........... stacked areas');
  WriteLn('  percent-stacked ... stacked by percentage');
  Halt;
end;

const
  FILE_NAME = 'area';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  dir, fn: String;
  stackMode: TsChartStackMode = csmDefault;
  rotated: Boolean = false;
  i: Integer;
begin
  fn := FILE_NAME;

  for i := 1 to ParamCount do
    case lowercase(ParamStr(i)) of
      'rotated':
        rotated := true;
      'normal' ,'default':
        stackMode := csmDefault;
      'stacked':
        stackMode := csmStacked;
      'percent-stacked', 'stacked-percent', 'percentstacked', 'stackedpercent', 'percent', 'percentage':
        stackMode := csmStackedPercentage;
      else
        WriteHelp;
    end;

  if rotated then
    fn := fn + '-rotated';
  case stackMode of
    csmDefault: ;
    csmStacked: fn := fn + '-stacked';
    csmStackedPercentage: fn := fn + '-stackedpercent';
  end;

  dir := ExtractFilePath(ParamStr(0)) + 'files/';
  ForceDirectories(dir);

  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('area_series');

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

    // Create chart: left/top in cell D4 of worksheet "area_series", 160 mm x 100 mm
    ch := sheet.AddChart(160, 100, 2, 3);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'School Grades';
    ch.Title.Font.Style := [fssBold];
    ch.Title.Font.Color := scBlue;
    ch.Legend.Border.Style := clsNoLine;
    ch.XAxis.Title.Caption := '';
    ch.YAxis.Title.Caption := 'Grade points';
    ch.YAxis.AxisLine.Color := ChartColor(scSilver);
    ch.YAxis.MajorTicks := [];
    ch.RotatedAxes := rotated;
    ch.StackMode := stackMode;

    // Add 1st area series ("Student 1")
    ser := TsAreaSeries.Create(ch);
    ser.SetTitleAddr(2, 1);              // series 1 title in cell B3
    ser.SetLabelRange(3, 0, 10, 0);      // series 1 x labels in A4:A11
    ser.SetYRange(3, 1, 10, 1);          // series 1 y values in B4:B11
    ser.Line.Color := ChartColor(scDarkRed);
    ser.Fill.Style := cfsSolidHatched;
    ser.Fill.Hatch := ch.Hatches.AddLineHatch('Crossed', chsDouble, ChartColor(scDarkRed), 2, 0.1, 45);
    ser.Fill.Color := ChartColor(scRed);

    // Add 2nd area series ("Student 2")
    ser := TsAreaSeries.Create(ch);
    ser.SetTitleAddr(2, 2);              // series 2 title in cell C3
    ser.SetLabelRange(3, 0, 10, 0);      // series 2 x labels in A4:A11
    ser.SetYRange(3, 2, 10, 2);          // series 2 y values in C4:C11
    ser.Line.Color := ChartColor(scDarkBlue);
    ser.Fill.Style := cfsSolidHatched;
    ser.Fill.Hatch := ch.Hatches.AddLineHatch('Forward', chsSingle, ChartColor(scWhite), 1.5, 0.1, 45);
    ser.Fill.Color := ChartColor(scBlue);

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

