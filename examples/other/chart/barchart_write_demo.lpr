program barchart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

procedure WriteHelp;
begin
  WriteLn('SYNTAX: barchart_write_demo [horz|vert] [side-by-side|stacked|percent-stacked] ');
  WriteLn('  vert .............. vertical bars');
  WriteLn('  horiz ............. horizontal bars');
  WriteLn('  side-by-side ...... bars side-by-side (default)');
  WriteLn('  stacked ........... stacked bars');
  WriteLn('  percent-stacked ... stacked by percentage');
  Halt;
end;

const
  FILE_NAME = 'bars';
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
  if ParamCount = 0 then
    WriteHelp;

  fn := FILE_NAME;

  for i := 1 to ParamCount do
    case lowercase(ParamStr(i)) of
      'hor', 'horiz', 'horizontal':
        rotated := true;
      'vert', 'vertical', 'rotated':
        rotated := false;
      'stacked':
        stackMode := csmStacked;
      'side-by-side':
        stackMode := csmDefault;
      'percent-stacked', 'stacked-percent', 'percentstacked', 'stackedpercent', 'percentage', 'percent':
        stackMode := csmStackedPercentage;
    end;

  case rotated of
    false: fn := fn + '-vert';
    true: fn := fn + '-horiz';
  end;
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
    sheet := book.AddWorksheet('bar_series');

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
    ch := book.AddChart(sheet, 2, 3, 160, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'School Grades';
    ch.Title.Font.Style := [fssBold];
    ch.Legend.Border.Style := clsNoLine;
    ch.XAxis.Title.Caption := '';
    ch.YAxis.Title.Caption := 'Grade points';
    ch.YAxis.AxisLine.Color := scSilver;
    ch.YAxis.MajorTicks := [];
    ch.RotatedAxes := rotated;
    ch.StackMode := stackMode;
    ch.BarGapWidthPercent := 75;

    // Add 1st bar series ("Student 1")
    ser := TsBarSeries.Create(ch);
    ser.SetTitleAddr(2, 1);              // series 1, title in cell B3
    ser.SetLabelRange(3, 0, 10, 0);      // series 1, x labels in A4:A11
    ser.SetYRange(3, 1, 10, 1);          // series 1, y values in B4:B11
    ser.Line.Color := scDarkRed;
    ser.Fill.Style := cfsSolidHatched;
    ser.Fill.Hatch := ch.Hatches.AddLineHatch('Crossed', chsDouble, scDarkRed, 2, 0.1, 45);
    ser.Fill.Color := scRed;
    ser.DataLabels := [cdlValue];        // Show scores as datapoint labels

    // Add 2nd bar series ("Student 2")
    ser := TsBarSeries.Create(ch);
    ser.SetTitleAddr(2, 2);              // series 2, title in cell C3
    ser.SetLabelRange(3, 0, 10, 0);      // series 2, x labels in A4:A11
    ser.SetYRange(3, 2, 10, 2);          // series 2, y values in C4:C11
    ser.Line.Color := scDarkBlue;
    ser.Fill.Style := cfsSolidHatched;
    ser.Fill.Hatch := ch.Hatches.AddLineHatch('Forward', chsSingle, scWhite, 1.5, 0.1, 45);
    ser.Fill.Color := scBlue;
    ser.DataLabels := [cdlValue];        // Show scores as datapoint labels

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

