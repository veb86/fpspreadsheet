program demo_richtext_utf8;

{$mode objfpc}{$H+}

uses
  SysUtils, fpstypes, fpspreadsheet, xlsxooxml, typinfo;

var
  book: TsWorkbook;
  sheet: TsWorksheet;
  cell: PCell;
  i: Integer;
  rtp: TsRichTextParam;
  fmt: TsCellFormat;
  dir: String;
begin
  dir := ExtractFilePath(ParamStr(0));

  book := TsWorkbook.Create;
  try
    // Prepare a worksheet containing rich-text in cell A1
    sheet := book.AddWorksheet('Sheet');
    sheet.WriteTextAsHtml(0, 0, 'äöü <b>ÄÖÜ</b> 123');

    // -----------------------------------------
    // Analyze the fonts used in cell A1
    //------------------------------------------

    cell := sheet.FindCell(0, 0);

    // Write a "ruler" for counting character positions
    WriteLn('12345678901234567890');
    // Write the unformatted cell text
    WriteLn(sheet.ReadAsText(cell));
    WriteLn;

    // characters before the first rich-text parameter have the cell font
    fmt := book.GetCellFormat(cell^.FormatIndex);  // get cell format record which contains the font index
    WriteLn(Format('Initial cell font: #%d (%s)', [fmt.FontIndex, book.GetFontAsString(fmt.FontIndex)]));

    // now write the rich-text parameters
    for rtp in cell^.RichTextParams do begin
      WriteLn(Format('Font #%d (%s) starting at character position %d', [
        rtp.FontIndex,
        book.GetFontAsString(rtp.FontIndex),
        rtp.FirstIndex
      ]));
    end;

    book.WriteToFile(dir+'test.xlsx', true);
  finally
    book.Free;
  end;

  {$IFDEF MSWindows}
  if ParamCount = 0 then
  begin
    WriteLn('Press ENTER to close...');
    ReadLn;
  end;
  {$ENDIF}
end.

