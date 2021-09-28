unit ViewerForm;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Grids,
  fpSpreadsheet, fpsTypes, fpsUtils;

type

  { TSpreadsheetViewerForm }

  TSpreadsheetViewerForm = class(TForm)
    StringGrid: TStringGrid;
    procedure FormCreate(Sender: TObject);
  private
  public
    procedure LoadFile(const AFileName, ASheetName: String);
  end;

var
  SpreadsheetViewerForm: TSpreadsheetViewerForm;

implementation

{$R *.lfm}

procedure TSpreadsheetViewerForm.FormCreate(Sender: TObject);
begin
  StringGrid.ColWidths[0] := 40;
end;

procedure TSpreadsheetViewerForm.LoadFile(const AFileName, ASheetName: String);
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  lastCol, lastRow: Integer;
  row, col: Integer;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.ReadFromFile(AFileName);
    worksheet := workbook.GetWorksheetByName(ASheetName);
    lastCol := worksheet.GetLastColIndex(true);
    lastRow := worksheet.GetLastRowIndex(true);
    StringGrid.RowCount := lastRow + 1 + StringGrid.FixedRows;
    StringGrid.ColCount := lastcol + 1 + StringGrid.FixedCols;
    for col := 1 to stringGrid.ColCount-1 do
      StringGrid.Cells[col, 0] := GetColString(col - 1);
    for row := 1 to StringGrid.RowCount-1 do
    begin
      StringGrid.Cells[0, row] := IntToStr(row);
      for col := 1 to StringGrid.ColCount - 1 do
        StringGrid.Cells[col, row] := worksheet.ReadAsText(row - 1, col - 1);
    end;
  finally
    workbook.Free;
  end;
end;

end.

