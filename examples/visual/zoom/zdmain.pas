unit zdMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Spin, fpstypes, fpspreadsheet, fpspreadsheetgrid, Types;

type

  { TMainForm }

  TMainForm = class(TForm)
    BtnOpen: TButton;
    CbOverrideZoomFactor: TCheckBox;
    edZoom: TSpinEdit;
    Grid: TsWorksheetGrid;
    OpenDialog: TOpenDialog;
    procedure BtnOpenClick(Sender: TObject);
    procedure edZoomEditingDone(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure GridMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
  private
    FWorkbook: TsWorkbook;
    procedure LoadFile(const AFileName: String);
    procedure UpdateCaption;

  public

  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

uses
  fpsRegFileFormats;


{ TMainForm }

procedure TMainForm.BtnOpenClick(Sender: TObject);
begin
 if FWorkbook <> nil then
   OpenDialog.InitialDir := ExtractFileDir(FWorkbook.FileName);
 if OpenDialog.Execute then
   LoadFile(OpenDialog.FileName);
end;

{ Set the zoom factor to the value in the edit control. Is called after
  pressing ENTER. }
procedure TMainForm.edZoomEditingDone(Sender: TObject);
begin
  Grid.ZoomFactor := edZoom.Value / 100;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
  priorityFormats: Array[0..7] of TsSpreadFormatID;
begin
  priorityFormats[0] := ord(sfOOXML);
  priorityFormats[1] := ord(sfExcel8);
  priorityFormats[2] := ord(sfExcel5);
  priorityFormats[3] := ord(sfExcel2);
  priorityFormats[4] := ord(sfExcelXML);
  priorityFormats[5] := ord(sfOpenDocument);
  priorityFormats[6] := ord(sfCSV);
  priorityFormats[7] := ord(sfHTML);

  OpenDialog.Filter := GetFileFormatFilter('|', ';', faRead, priorityFormats, true, true);
end;

{ Mouse wheel event handler for setting the zoom factor using the mouse wheel
  (together with pressing the SHFT + CTRL keys) }
procedure TMainForm.GridMouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
  if ([ssCtrl, ssShift] * Shift = [ssCtrl, ssShift]) then begin
    if WheelDelta > 0 then
      Grid.ZoomFactor := Grid.ZoomFactor * 1.05
    else
      Grid.ZoomFactor := Grid.ZoomFactor / 1.05;
    edZoom.Value := round(Grid.ZoomFactor * 100);
    Handled := true;
  end;
end;

procedure TMainForm.LoadFile(const AFileName: String);
var
  crs: TCursor;
  book: TsWorkbook;
begin
  crs := Screen.Cursor;
  try
    if CbOverrideZoomFactor.Checked then begin
      book := TsWorkbook.Create;
      try
        Screen.Cursor := crHourglass;
        // Read the file
        book.ReadFromFile(AFilename);
        // If you want to override the zoom factor of the file set it before
        // assigning the worksheet to the grid.
        book.GetFirstWorksheet.ZoomFactor := edZoom.Value / 100;
        // Load the worksheet into the grid.
        Grid.LoadFromWorkbook(book, 0);
      except
        on E:Exception do begin
          MessageDlg(Format('File "%s" cannot be loaded.'#13 + E.Message, [AFilename]),
            mtError, [mbOK], 0);
          book.Free;
        end;
      end;
    end else begin
      Grid.LoadSheetFromSpreadsheetFile(AFilename, 0);
      edZoom.Value := Grid.ZoomFactor * 100;
    end;
    UpdateCaption;
  finally
    Screen.Cursor := crs;
  end;
end;

procedure TMainForm.UpdateCaption;
begin
  if FWorkbook = nil then
    Caption := 'Zoom demo'
  else
    Caption := 'Zoom demo - "' + FWorkbook.Filename + '"';
end;

end.

