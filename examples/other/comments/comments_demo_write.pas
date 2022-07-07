program comments_demo_write;

uses
  fpspreadsheet, fpstypes, fpsallformats;

const
  FILE_NAME = 'test';
  
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  
begin
  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Sheet 1');
    worksheet.WriteText(0, 0, 'Angle');
    worksheet.WriteNumber(0, 1, 30.0, nfFixed, 1);
    worksheet.WriteComment(0, 1, 'Enter angle in degrees here.');
    worksheet.WriteText(1, 0, 'sin(Angle)');
    worksheet.WriteFormula(1, 1, '=sin(B1*pi()/180)');
    workbook.WriteToFile(FILE_NAME + '.xlsx', sfOOXML, true);
//    workbook.WriteToFile(FILE_NAME + '8.xls', sfExcel8, true);  // no BIFF8 writing support for comments so far.
    workbook.WriteToFile(FILE_NAME + '5.xls', sfExcel5, true);
    workbook.WriteToFile(FILE_NAME + '2.xls', sfExcel2, true);
    workbook.WriteToFile(FILE_NAME + '.ods', sfOpenDocument, true);
  finally
    workbook.Free;
  end;
end.

