program stock_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'stock';

var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsStockSeries;
  r: Integer;
  d: TDate;
  dir, fn: String;
  candlestickMode: Boolean;
  rotated: Boolean;

  procedure WriteHelp;
  begin
    WriteLn('SYNTAX: stock_write_demo hlc|candlestick [rotated]');
    WriteLn('  hlc ........... Create high-low-close series');
    WriteLn('  candlestick ... Create candle-stick series');
    WriteLn('  rotated ....... optional: rotated axes (date vertical)');
    halt;
  end;

  procedure WriteData(var ARow: Integer; var ADate: TDate; AOpen, AHigh, ALow, AClose: Double);
  begin
    sheet.WriteDateTime(ARow, 0, ADate, nfShortDate);
    sheet.WriteNumber  (ARow, 1, AOpen, nfFixed, 2);
    sheet.WriteNumber  (ARow, 2, AHigh, nfFixed, 2);
    sheet.WriteNumber  (ARow, 3, ALow,  nfFixed, 2);
    sheet.WriteNumber  (ARow, 4, AClose,nfFixed, 2);
    inc(ARow);
    ADate := ADate + 1;
  end;

begin
  if ParamCount >= 1 then
  begin
    case lowercase(ParamStr(1)) of
      'hlc':
        begin
          candleStickMode := false;
          fn := FILE_NAME + '-hlc';
        end;
      'candlestick':
        begin
          candleStickMode := true;
          fn := FILE_NAME + '-candlestick';
        end;
      else
        WriteHelp;
    end;
    rotated := (ParamCount >= 2) and (lowercase(ParamStr(2)) = 'rotated');
    if rotated then fn := fn + '-rotated';
  end else
    WriteHelp;

  dir := ExtractFilePath(ParamStr(0)) + 'files/';
  ForceDirectories(dir);

  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('test');

    // Enter data
    sheet.WriteText  (0, 0, 'My Company');
    sheet.WriteFont  (0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText  ( 2, 0, 'Date');
    sheet.WriteText  ( 2, 1, 'Open');
    sheet.WriteText  ( 2, 2, 'High');
    sheet.WriteText  ( 2, 3, 'Low');
    sheet.WriteText  ( 2, 4, 'Close');
    d := EncodeDate(2023, 3, 6);
    r := 3;      //  O    H    L    C
    WriteData(r, d, 100, 110,  95, 105);
    WriteData(r, d, 107, 112, 101, 104);
    WriteData(r, d, 108, 113, 100, 106);
    WriteData(r, d, 109, 115,  99, 110);
    WriteData(r, d, 110, 119, 103, 115);

    // Create chart: left/top in cell D4, 160 mm x 100 mm
    ch := book.AddChart(sheet, 2, 5, 160, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;
    ch.RotatedAxes := rotated;
    ch.XAxis.DateTime := true;
    ch.XAxis.Title.Caption := 'Date';
    ch.XAxis.MajorGridLines.Style := clsNoLine;
    ch.XAxis.MinorGridLines.Style := clsNoLine;
    ch.YAxis.Title.Caption := 'Stock price';
    ch.YAxis.MajorGridLines.Style := clsSOLID;
    ch.YAxis.MinorGridLines.Style := clsNoLine;
    ch.YAxis.AutomaticMin := false;
    ch.YAxis.AutomaticMax := false;
    ch.YAxis.Max := 120;
    ch.YAxis.Min := 90;

    // Add stock series
    ser := TsStockSeries.Create(ch);

    // Series properties
    ser.CandleStick := candleStickMode;
    ser.CandleStickUpFill.Color := scGreen;
    ser.CandlestickDownFill.Color := scRed;
    ser.SetTitleAddr (0, 0);
    if candleStickMode then ser.SetOpenRange (3, 1, 7, 1);
    ser.SetHighRange (3, 2, 7, 2);
    ser.SetLowRange  (3, 3, 7, 3);
    ser.SetCloseRange(3, 4, 7, 4);
    ser.SetXRange    (3, 0, 7, 0);
    ser.SetLabelRange(3, 0, 7, 0);
    ser.RangeLine.Width := 1;
    ser.RangeLine.Color := scRed;

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

