program errorbars_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'errorbars';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsScatterSeries;
  dir, fn: String;
  errorRange: Boolean = false;
begin
  // Error bar kind for y bars only. x bars are always constant
  if (ParamCount > 0) and (lowercase(ParamStr(1)) = 'range') then
    errorRange := true;

  fn := FILE_NAME;
  if errorRange then
    fn := fn + '-range'
  else
    fn := fn + '-percentage';

  dir := ExtractFilePath(ParamStr(0)) + 'files/';
  ForceDirectories(dir);

  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('errorbars_test');

    // Enter data
    sheet.WriteText(0, 0, 'Data');
    sheet.WriteFont(0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText  (2, 0, 'x');  sheet.Writetext  (2, 1, 'y');   sheet.WriteText  (2, 2, 'dy');
    sheet.WriteNumber(3, 0, 1.1);  sheet.WriteNumber(3, 1,  9.0);  sheet.WriteNumber(3, 2, 0.5);
    sheet.WriteNumber(4, 0, 1.9);  sheet.WriteNumber(4, 1, 20.5);  sheet.WriteNumber(4, 2, 3.5);
    sheet.WriteNumber(5, 0, 2.5);  sheet.WriteNumber(5, 1, 24.5);  sheet.WriteNumber(5, 2, 2.7);
    sheet.WriteNumber(6, 0, 3.1);  sheet.WriteNumber(6, 1, 33.2);  sheet.WriteNumber(6, 2, 3.1);
    sheet.WriteNumber(7, 0, 5.2);  sheet.WriteNumber(7, 1, 49.4);  sheet.WriteNumber(7, 2, 6.7);
    sheet.WriteNumber(8, 0, 6.8);  sheet.WriteNumber(8, 1, 71.3);  sheet.WriteNumber(8, 2, 3.5);

    // Create chart: left/top in cell D4, 150 mm x 100 mm
    ch := book.AddChart(sheet, 2, 3, 150, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;

    // Add scatter series
    ser := TsScatterSeries.Create(ch);

    // Series properties
    ser.SetTitleAddr(0, 0);
    ser.SetXRange(3, 0, 8, 0);
    ser.SetYRange(3, 1, 8, 1);
    ser.ShowLines := false;
    ser.ShowSymbols := true;
    ser.Symbol := cssCircle;
    ser.SymbolFill.Style := cfsSolid;
    ser.SymbolFill.Color := scRed;
    ser.SymbolBorder.Style := clsNoLine;

    ser.XErrorBars.Visible := true;
    ser.XErrorBars.Kind := cebkConstant;
    ser.XErrorBars.ValuePos := 0.5;
    ser.XErrorBars.ValueNeg := 0.5;
    ser.XErrorBars.Line.Color := scRed;

    ser.YErrorBars.Visible := true;
    if errorRange then
    begin
      ser.YErrorBars.Kind := cebkCellRange;
      ser.YErrorBars.SetErrorBarRangePos(3, 2, 8, 2);
      ser.YErrorBars.SetErrorBarRangeNeg(3, 2, 8, 2);
    end else
    begin
      ser.YErrorBars.Kind := cebkPercentage;
      ser.YErrorBars.ValuePos := 10;  // percent
      ser.YErrorBars.ValueNeg := 10;  // percent
    end;
    ser.YErrorBars.Line.Color := scRed;

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

