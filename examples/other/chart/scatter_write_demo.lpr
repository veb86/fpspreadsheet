program scatter_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

procedure WriteHelp;
begin
  WriteLn('SYNTAX: scatter_write_demo lin|log|loglog');
  WriteLn('  lin ........... Both axes linear');
  WriteLn('  log ........... y axis logarithmic');
  WriteLn('  loglog ........ Both axes logarithmic');
  halt;
end;

const
  FILE_NAME = 'scatter';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsScatterSeries;
  fn: String;
  mode: Integer;  // 0=linear, 1=log, 2=log-log

begin
  if ParamCount >= 1 then
  begin
    case lowercase(ParamStr(1)) of
      'lin':
        begin
          mode := 0;
          fn := FILE_NAME + '-lin';
        end;
      'log':
        begin
          mode := 1;
          fn := FILE_NAME + '-log';
        end;
      'loglog', 'log-log':
        begin
          mode := 2;
          fn := FILE_NAME + '-loglog';
        end;
      else
        WriteHelp;
    end;
  end else
    WriteHelp;

  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('test');

    // Enter data
    sheet.WriteText(0, 0, 'Data');
    sheet.WriteFont(0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText         ( 2, 0,  'x');   sheet.WriteText   ( 2, 1, 'y');
    case mode of
      0: begin
           sheet.WriteNumber( 3, 0,  0.1);  sheet.WriteFormula( 3, 1, 'A4^2');
           sheet.WriteNumber( 4, 0,  8.8);  sheet.WriteFormula( 4, 1, 'A5^2');
           sheet.WriteNumber( 5, 0, 16.9);  sheet.WriteFormula( 5, 1, 'A6^2');
           sheet.WriteNumber( 6, 0, 24.6);  sheet.WriteFormula( 6, 1, 'A7^2');
           sheet.WriteNumber( 7, 0, 38.3);  sheet.WriteFormula( 7, 1, 'A8^2');
           sheet.WriteNumber( 8, 0, 45.9);  sheet.WriteFormula( 8, 1, 'A9^2');
           sheet.WriteNumber( 9, 0, 55.6);  sheet.WriteFormula( 9, 1, 'A10^2');
           sheet.WriteNumber(10, 0, 68.3);  sheet.WriteFormula(10, 1, 'A11^2');
         end;
      1: begin
           sheet.WriteNumber(3, 0, 0.1);  sheet.WriteFormula(3, 1, 'exp(A4)');
           sheet.WriteNumber(4, 0, 0.8);  sheet.WriteFormula(4, 1, 'exp(A5)');
           sheet.WriteNumber(5, 0, 1.4);  sheet.WriteFormula(5, 1, 'exp(A6)');
           sheet.WriteNumber(6, 0, 2.6);  sheet.WriteFormula(6, 1, 'exp(A7)');
           sheet.WriteNumber(7, 0, 4.3);  sheet.WriteFormula(7, 1, 'exp(A8)');
           sheet.WriteNumber(8, 0, 5.9);  sheet.WriteFormula(8, 1, 'exp(A9)');
           sheet.WriteNumber(9, 0, 7.5);  sheet.WriteFormula(9, 1, 'exp(A10)');
         end;
      2: begin
           sheet.WriteNumber(3, 0, 0.1);  sheet.WriteFormula(3, 1, 'A4^2');
           sheet.WriteNumber(4, 0, 0.8);  sheet.WriteFormula(4, 1, 'A5^2');
           sheet.WriteNumber(5, 0, 1.9);  sheet.WriteFormula(5, 1, 'A6^2');
           sheet.WriteNumber(6, 0, 4.6);  sheet.WriteFormula(6, 1, 'A7^2');
           sheet.WriteNumber(7, 0, 8.3);  sheet.WriteFormula(7, 1, 'A8^2');
           sheet.WriteNumber(8, 0,15.9);  sheet.WriteFormula(8, 1, 'A9^2');
           sheet.WriteNumber(9, 0,25.6);  sheet.WriteFormula(9, 1, 'A10^2');
           sheet.WriteNumber(10,0,68.3);  sheet.WriteFormula(10,1, 'A11^2');
         end;
    end;

    // Create chart: left/top in cell D4, 150 mm x 100 mm
    ch := book.AddChart(sheet, 2, 3, 150, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;
    // Set up logarithmic axes if needed.
    case mode of
      0: ;
      1: ch.YAxis.Logarithmic := true;
      2: begin
           ch.XAxis.Logarithmic := true;
           ch.XAxis.Max := 100;
           ch.XAxis.AutomaticMax := false;
           ch.YAxis.Logarithmic := true;
        end;
    end;

    // Add scatter series
    ser := TsScatterSeries.Create(ch);

    // Series properties
    ser.SetTitleAddr(0, 0);
    ser.SetXRange(3, 0, 10, 0);    // A4:A11
    ser.SetYRange(3, 1, 10, 1);    // B4:B11
    ser.ShowLines := true;
    ser.ShowSymbols := true;
    ser.Symbol := cssCircle;
    ser.SymbolBorder.Style := clsNoLine;
//    ser.Line.Style := clsDash;
    ser.Line.Width := 0.5;  // mm

    book.WriteToFile(fn + '.xlsx', true);
    WriteLn('Data saved with chart to ', fn + '.xlsx');

    book.Options := book.Options + [boCalcBeforeSaving];
    book.WriteToFile(fn + '.ods', true);
    WriteLn('Data saved with chart to ', fn + '.ods');
  finally
    book.Free;
  end;
end.

