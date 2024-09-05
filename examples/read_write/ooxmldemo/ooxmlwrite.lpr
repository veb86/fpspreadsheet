{
ooxmlwrite.lpr

Demonstrates how to write an OOXML file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program ooxmlwrite;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, xlsxOOXML, fpscell;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  i: Integer;
  MyCell: PCell;

begin
  // Open the output file
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  // Write some number cells
  MyWorksheet.WriteNumber(0, 0, 1.0);
  MyWorksheet.WriteNumber(0, 1, 2.0);
  MyWorksheet.WriteNumber(0, 2, 3.0);
  MyWorksheet.WriteNumber(0, 3, 4.0);

  // Write text with special xml characters
  MyWorksheet.WriteText(0, 4, '& " '' < >');

  MyWorksheet.WriteText(0, 26, 'AA'); // Test for column name

  MyWorksheet.WriteColWidth(0, 20, suChars);
  MyWorksheet.WriteRowHeight(0, 4, suLines);

  // Write some formulas
  Myworksheet.WriteFormula(0, 5, '=A1-B1');
  Myworksheet.WriteFormula(0, 6, '=SUM(A1:D1)');
  MyWorksheet.WriteFormula(0, 7, '=SIN(A1+B1)');

  // Test for built-in bold
  MyCell := MyWorksheet.GetCell(2, 0);
  MyCell^.FontIndex := BOLD_FONTINDEX;
  MyWorksheet.WriteText(MyCell,'Bold');

  // Test for built-in italic
  MyCell := MyWorksheet.GetCell(2, 1);
  MyCell^.FontIndex := ITALIC_FONTINDEX;
  Myworksheet.WriteText(MyCell, 'Italic');

  // Test for bold-italic
  MyCell := MyWorksheet.GetCell(2, 2);
  MyWorksheet.WriteFontStyle(MyCell, [fssBold, fssItalic]);
  MyWorksheet.WriteText(MyCell, 'Bold-Italic');

  // Test for underline
  MyCell := MyWorksheet.WriteText(2, 3, 'Underlined');
  MyWorksheet.WriteFontStyle(MyCell, [fssUnderline]);

  // Test for strike-through
  MyCell := MyWorksheet.WriteText(2, 4, 'Strike-out');
  MyWorksheet.WriteFontStyle(MyCell, [fssStrikeOut]);

  // Background and text color
  MyWorksheet.WriteText(4, 0, 'white on red');
  Myworksheet.WriteBackgroundColor(4, 0, scRed);
  MyWorksheet.WriteFontColor(4, 0, scWhite);

  // Border
  MyWorksheet.WriteText(4, 2, 'left/right border');
  Myworksheet.WriteBorders(4, 2, [cbWest, cbEast]);
  MyWorksheet.WriteHorAlignment(4, 2, haCenter);
  MyWorksheet.WriteWordWrap(4, 2, true);

  Myworksheet.WriteText(4, 4, 'top/bottom border');
  Myworksheet.WriteBorders(4, 4, [cbNorth, cbSouth]);
  MyWorksheet.WriteBorderStyle(4, 4, cbSouth, lsThick, scBlue);
  Myworksheet.WriteHorAlignment(4, 4, haRight);
  MyWorksheet.WriteWordwrap(4, 4, true);

  // Wordwrap
  MyWorksheet.WriteText(4, 6, 'This is a long, long, long, wrapped text.');
  MyWorksheet.WriteWordwrap(4, 6, true);

  // Write name of this binary
  MyWorksheet.WriteText(6, 0, ParamStr(0));

    // Create a new worksheet
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet 2');

  // Write some string cells
  MyWorksheet.WriteText(0, 0, 'First');
  MyWorksheet.WriteText(0, 1, 'Second');
  MyWorksheet.WriteText(0, 2, 'Third');
  MyWorksheet.WriteText(0, 3, 'Fourth');

  // Write current date/time
  MyWorksheet.WriteDateTime(0, 5, now, nfShortDate);
  MyWorksheet.WriteDateTime(1, 5, now, nfShortTime);
  MyWorksheet.WriteDateTime(2, 5, now, 'nn:ss.zzz');

  // Write some numbers in various formats
  MyWorksheet.WriteNumber  (0, 6, 12345.6789, nfFixed, 0);
  MyWorksheet.WriteNumber  (1, 6, 12345.6789, nfFixed, 3);
  MyWorksheet.WriteNumber  (2, 6, 12345.6789, nfFixedTh, 0);
  MyWorksheet.Writenumber  (3, 6, 12345.6789, nfFixedTh, 3);
  MyWorksheet.WriteNumber  (4, 6, 12345.6789, nfExp, 2);
  Myworksheet.Writenumber  (5, 6, 12345.6789, nfExp, 4);
  MyWorksheet.WriteCurrency(6, 6,-12345.6789, nfCurrency, 2);
  MyWorksheet.WriteCurrency(7, 6,-12345.6789, nfCurrencyRed, 2);
  MyWorksheet.WriteNumber  (8, 6, 1.66666667, nfFraction, '# ?/?');

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.xlsx', sfOOXML, true);
  MyWorkbook.Free;

  WriteLn('Workbook written to "' + Mydir + 'test.xlsx' + '".');

  {$IFDEF MSWINDOWS}
  WriteLn('Press ENTER to quit...');
  ReadLn;
  {$ENDIF}
end.

