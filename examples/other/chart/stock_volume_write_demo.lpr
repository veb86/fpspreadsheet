program stock_volume_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'stock-vol';

var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsStockSeries;
  vser: TsChartSeries;
  r: Integer;
  d: TDate;
  fn: String;
  candleStickMode: Boolean;
  volumeMode: char;
  rotated: Boolean;

  procedure WriteHelp;
  begin
    WriteLn('SYNTAX: stock_volume_write_demo hlc|candlestick bar|area|line [rotated]');
    WriteLn('  hlc ........... Create high-low-close series');
    WriteLn('  candlestick ... Create candle-stick series');
    WriteLn('  area .......... Display volume as area series');
    WriteLn('  bar ........... Display volume as bar series');
    WriteLn('  line .......... Display volume as line series');
    WriteLn('  rotated ....... (optional) rotated axes (date axis vertical)');
    halt;
  end;

  procedure WriteData(var ARow: Integer; var ADate: TDate; AVolume, AOpen, AHigh, ALow, AClose: Double);
  begin
    sheet.WriteDateTime(ARow, 0, ADate,   nfShortDate);
    sheet.WriteNumber  (ARow, 1, AVolume, nfFixed, 0);
    sheet.WriteNumber  (ARow, 2, AOpen,   nfFixed, 2);
    sheet.WriteNumber  (ARow, 3, AHigh,   nfFixed, 2);
    sheet.WriteNumber  (ARow, 4, ALow,    nfFixed, 2);
    sheet.WriteNumber  (ARow, 5, AClose,  nfFixed, 2);
    inc(ARow);
    ADate := ADate + 1;
  end;

begin
  if ParamCount >= 2 then
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
    case lowercase(ParamStr(2)) of
      'area':
        begin
          volumeMode := 'a';
          fn := fn + '-area';
        end;
      'bar', 'bars':
        begin
          volumeMode := 'b';
          fn := fn + '-bars';
        end;
      'line':
        begin
          volumeMode := 'l';
          fn := fn + '-line';
        end;
      else
        WriteHelp;
    end;
    rotated := (ParamCount >= 3) and (lowercase(ParamStr(3)) = 'rotated');
    if rotated then fn := fn + '-rotated';
  end else
    WriteHelp;

  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('test');

    // Enter data
    sheet.WriteText  (0, 0, 'My Company');
    sheet.WriteFont  (0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText  ( 2, 0, 'Date');
    sheet.WriteText  ( 2, 1, 'Volume');
    sheet.WriteText  ( 2, 2, 'Open');
    sheet.WriteText  ( 2, 3, 'High');
    sheet.WriteText  ( 2, 4, 'Low');
    sheet.WriteText  ( 2, 5, 'Close');
    d := EncodeDate(2023, 3, 6);
    r := 3;       //  Vol    O    H    L    C
    WriteData(r, d, 100000, 100, 110,  95, 105);
    WriteData(r, d,  90000, 107, 112, 101, 104);
    WriteData(r, d, 120000, 108, 113, 100, 106);
    WriteData(r, d, 110000, 109, 115,  99, 110);
    WriteData(r, d,  95000, 110, 119, 103, 115);

    // Create chart: left/top in cell D4, 160 mm x 100 mm
    ch := book.AddChart(sheet, 2, 6, 160, 100);

    // Chart properties
    ch.RotatedAxes := rotated;
    
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;
    ch.Legend.Position := lpBottom;

    ch.XAxis.DateTime := true;
    ch.XAxis.Title.Caption := 'Date';
    ch.XAxis.MajorGridLines.Style := clsNoLine;
    ch.XAxis.MinorGridLines.Style := clsNoLine;

    ch.YAxis.Title.Caption := 'Stock price';
    ch.YAxis.MajorGridLines.Style := clsSolid;
    ch.YAxis.MinorGridLines.Style := clsNoLine;
    ch.YAxis.AutomaticMin := false;
    ch.YAxis.AutomaticMax := false;
    ch.YAxis.Max := 120;
    ch.YAxis.Min := 80;

    ch.Y2Axis.Title.Caption := 'Volume';
    ch.Y2Axis.MajorGridLines.Style := clsNoLine;
    ch.Y2Axis.MinorGridLines.Style := clsNoLine;
    ch.Y2Axis.AutomaticMax := false;
    ch.Y2Axis.AutomaticMax := false;
    ch.Y2Axis.Min := 0;
    ch.Y2Axis.Max := 300000;

    // Add stock series
    ser := TsStockSeries.Create(ch);

    // Stock series properties
    ser.YAxis := calPrimary;
    ser.CandleStick := candleStickMode;
    ser.CandleStickUpFill.Color := scGreen;
    ser.CandlestickDownFill.Color := scRed;
    ser.SetTitleAddr (0, 0);
    if candleStickMode then ser.SetOpenRange (3, 2, 7, 2);
    ser.SetHighRange (3, 3, 7, 3);
    ser.SetLowRange  (3, 4, 7, 4);
    ser.SetCloseRange(3, 5, 7, 5);
    ser.SetLabelRange(3, 0, 7, 0);

    // Add series for volume data, type depending on 2nd commandline argument
    case volumeMode of
      'a': vser := TsAreaSeries.Create(ch);
      'b': vser := TsBarSeries.Create(ch);
      'l': vser := TsLineSeries.Create(ch);
    end;

    // Volume series properties
    vser.YAxis := calSecondary;
    vser.SetLabelRange(3, 0, 7, 0);
    vser.SetYRange    (3, 1, 7, 1);
    vser.SetTitleAddr (2, 1);

    {
    book.WriteToFile(fn + '.xlsx', true);   // Excel fails to open the file
    WriteLn('Data saved with chart to ', fn, '.xlsx');
    }

    book.WriteToFile(fn + '.ods', true);
    WriteLn('Data saved with chart to ', fn, '.ods');
  finally
    book.Free;
  end;
end.

