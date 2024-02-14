program linechart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

procedure WriteHelp;
begin
  WriteLn('SYNTAX: linechart_write_demo [rotated] [normal|stacked|percent-stacked] ');
  WriteLn('  (no argument) ..... normal orientation (x horizontal)');
  WriteLn('  rotated ........... axes rotated (x vertical)');
  WriteLn('  normal ............ back-to-bottom areas (default)');
  WriteLn('  stacked ........... stacked areas');
  WriteLn('  percent-stacked ... stacked by percentage');
  Halt;
end;

const
  FILE_NAME = 'line';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsLineSeries;
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
    sheet := book.AddWorksheet('line_series');

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
    ch.Title.Font.Color := scBlue;
    ch.Legend.Border.Style := clsNoLine;
    ch.XAxis.Title.Caption := '';
    ch.YAxis.Title.Caption := 'Grade points';
    ch.YAxis.AxisLine.Color := scSilver;
    ch.YAxis.MajorTicks := [];
    ch.RotatedAxes := rotated;
    ch.StackMode := stackMode;

    // Add 1st line series ("Student 1")
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(2, 1);              // series 1 title in cell B3
    ser.SetLabelRange(3, 0, 10, 0);      // series 1 x labels in A4:A11
    ser.SetYRange(3, 1, 10, 1);          // series 1 y values in B4:B11
    ser.Line.Color := scRed;
    ser.ShowSymbols := true;
    ser.SymbolFill.Color := scRed;
    ser.SymbolFill.Style := cfsSolid;
    ser.SymbolBorder.Color := scBlack;
    ser.Smooth := true;
//    ser.GroupIndex := -1;

    // Add 2nd line series ("Student 2")
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(2, 2);              // series 2 title in cell C3
    ser.SetLabelRange(3, 0, 10, 0);      // series 2 x labels in A4:A11
    ser.SetYRange(3, 2, 10, 2);          // series 2 y values in C4:C11
    ser.Line.Color := scBlue;
    ser.SymbolFill.Color := scBlue;
    ser.SymbolFill.Style := cfsSolid;
    ser.SymbolBorder.Color := scBlack;
    //ser.Smooth := true;
    ser.ShowSymbols := true;
//    ser.GroupIndex := -1;

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

