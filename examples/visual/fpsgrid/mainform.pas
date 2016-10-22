unit mainform; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, Menus, ExtCtrls, ActnList, Spin, Buttons, ButtonPanel,
  fpspreadsheetgrid, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    BtnOpen: TButton;
    BtnSave: TButton;
    BtnNew: TButton;
    CbReadFormulas: TCheckBox;
    CbAutoCalc: TCheckBox;
    Panel3: TPanel;
    SheetsCombo: TComboBox;
    Label1: TLabel;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    SaveDialog: TSaveDialog;
    WorksheetGrid: TsWorksheetGrid;
    procedure BtnNewClick(Sender: TObject);
    procedure BtnOpenClick(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure CbAutoCalcChange(Sender: TObject);
    procedure CbReadFormulasChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SheetsComboSelect(Sender: TObject);
  private
    { private declarations }
    procedure LoadFile(const AFileName: String);
  public
    { public declarations }
  end; 

var
  Form1: TForm1; 

implementation

uses
  fpcanvas, lazutf8,
  fpstypes, fpsutils, fpsReaderWriter, fpspreadsheet;


{ TForm1 }

procedure TForm1.BtnNewClick(Sender: TObject);
var
  dlg: TForm;
  edCols, edRows: TSpinEdit;
begin
  dlg := TForm.Create(nil);
  try
    dlg.Width := 220;
    dlg.Height := 128;
    dlg.Position := poMainFormCenter;
    dlg.Caption := 'New workbook';
    edCols := TSpinEdit.Create(dlg);
    with edCols do begin
      Parent := dlg;
      Left := dlg.ClientWidth - Width - 24;
      Top := 16;
      Value := WorksheetGrid.ColCount - ord(WorksheetGrid.ShowHeaders);
    end;
    with TLabel.Create(dlg) do begin
      Parent := dlg;
      Left := 24;
      Top := edCols.Top + 3;
      Caption := 'Columns:';
      FocusControl := edCols;
    end;
    edRows := TSpinEdit.Create(dlg);
    with edRows do begin
      Parent := dlg;
      Left := edCols.Left;
      Top := edCols.Top + edCols.Height + 8;
      Value := WorksheetGrid.RowCount - ord(WorksheetGrid.ShowHeaders);
    end;
    with TLabel.Create(dlg) do begin
      Parent := dlg;
      Left := 24;
      Top := edRows.Top + 3;
      Caption := 'Rows:';
      FocusControl := edRows;
    end;
    with TButtonPanel.Create(dlg) do begin
      Parent := dlg;
      Align := alBottom;
      ShowButtons := [pbCancel, pbOK];
    end;
    if dlg.ShowModal = mrOK then begin
      WorksheetGrid.NewWorkbook(edCols.Value, edRows.Value);
      SheetsCombo.Items.Clear;
      SheetsCombo.Items.Add('Sheet 1');
      SheetsCombo.ItemIndex := 0;
    end;
  finally
    dlg.Free;
  end;
end;

procedure TForm1.BtnOpenClick(Sender: TObject);
begin
  if OpenDialog.FileName <> '' then begin
    OpenDialog.InitialDir := ExtractFileDir(OpenDialog.FileName);
    OpenDialog.FileName := ChangeFileExt(ExtractFileName(OpenDialog.FileName), '');
  end;
  if OpenDialog.Execute then begin
    LoadFile(OpenDialog.FileName);
  end;
end;

// Saves sheet in grid to file, overwriting existing file
procedure TForm1.BtnSaveClick(Sender: TObject);
var
  err, fn: String;
begin
  if WorksheetGrid.Workbook = nil then
    exit;

  if WorksheetGrid.Workbook.Filename <>'' then begin
    fn := AnsiToUTF8(WorksheetGrid.Workbook.Filename);
    SaveDialog.InitialDir := ExtractFileDir(fn);
    SaveDialog.FileName := ChangeFileExt(ExtractFileName(fn), '');
  end;

  if SaveDialog.Execute then
  begin
    Screen.Cursor := crHourglass;
    try
      WorksheetGrid.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));
      //WorksheetGrid.WorkbookSource.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));     // works as well
    finally
      Screen.Cursor := crDefault;
      // Show a message in case of error(s)
      err := WorksheetGrid.Workbook.ErrorMsg;
      if err <> '' then
        MessageDlg(err, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TForm1.CbAutoCalcChange(Sender: TObject);
begin
  WorksheetGrid.AutoCalc := CbAutoCalc.Checked;
end;

procedure TForm1.CbReadFormulasChange(Sender: TObject);
begin
  WorksheetGrid.ReadFormulas := CbReadFormulas.Checked;
end;

procedure TForm1.FormCreate(Sender: TObject);
const
  THICK_BORDER: TsCellBorderStyle = (LineStyle: lsThick; Color: clNavy);
  MEDIUM_BORDER: TsCellBorderSTyle = (LineStyle: lsMedium; Color: clRed);
  DOTTED_BORDER: TsCellBorderSTyle = (LineStyle: lsDotted; Color: clRed);
begin
  // Add some cells and formats
  WorksheetGrid.ColWidths[1] := 180;
  WorksheetGrid.ColWidths[2] := 100;

  WorksheetGrid.Cells[1,1] := 'This is a demo';
  WorksheetGrid.MergeCells(1,1, 2,1);
  WorksheetGrid.HorAlignment[1,1] := haCenter;
  WorksheetGrid.CellBorders[1,1, 2,1] := [cbSouth];
  WorksheetGrid.CellBorderStyles[1,1, 2,1, cbSouth] := THICK_BORDER;
  WorksheetGrid.BackgroundColors[1,1, 2,1] := RGBToColor(220, 220, 220);
  WorksheetGrid.CellFontColor[1,1] := clNavy;
  WorksheetGrid.CellFontStyle[1,1] := [fssBold];

  WorksheetGrid.Cells[1,2] := 'Number:';
  WorksheetGrid.HorAlignment[1,2] := haRight;
  WorksheetGrid.CellFontStyle[1,2] := [fssItalic];
  WorksheetGrid.CellFontColor[1,2] := clNavy;
  WorksheetGrid.Cells[2,2] := 1.234;

  WorksheetGrid.Cells[1,3] := 'Date:';
  WorksheetGrid.HorAlignment[1,3] := haRight;
  WorksheetGrid.CellFontStyle[1,3] := [fssItalic];
  WorksheetGrid.CellFontColor[1,3] := clNavy;
  WorksheetGrid.NumberFormat[2,3] := 'mmm dd, yyyy';
  WorksheetGrid.Cells[2,3] := date;

  WorksheetGrid.Cells[1,4] := 'Time:';
  WorksheetGrid.HorAlignment[1,4] := haRight;
  WorksheetGrid.CellFontStyle[1,4] := [fssItalic];
  WorksheetGrid.CellFontColor[1,4] := clNavy;
  WorksheetGrid.NumberFormat[2,4] := 'hh:nn';
  WorksheetGrid.Cells[2,4] := now();

  WorksheetGrid.Cells[1,5] := 'Rich text:';
  WorksheetGrid.HorAlignment[1,5] := haRight;
  WorksheetGrid.CellFontStyle[1,5] := [fssItalic];
  WorksheetGrid.CellFontColor[1,5] := clNavy;
  WorksheetGrid.Cells[2,5] := '100 cm<sup>2</sup>';

  WorksheetGrid.Cells[1,6] := 'Formula:';
  WorksheetGrid.HorAlignment[1,6] := haRight;
  WorksheetGrid.CellFontStyle[1,6] := [fssItalic];
  WorksheetGrid.CellFontColor[1,6] := clNavy;
  WorksheetGrid.Cells[2,6] := '=B2^2*PI()';
  WorksheetGrid.CellComment[2,6] := 'Area of the circle with radius given in cell B2';
  WorksheetGrid.NumberFormat[2,6] := '0.000';

  CbAutoCalc.Checked := WorksheetGrid.AutoCalc;
  CbReadFormulas.Checked := WorksheetGrid.ReadFormulas;
end;

procedure TForm1.SheetsComboSelect(Sender: TObject);
begin
  WorksheetGrid.SelectSheetByIndex(SheetsCombo.ItemIndex);
end;

// Loads first worksheet from file into grid
procedure TForm1.LoadFile(const AFileName: String);
var
  err: String;
begin
  // Load file
  Screen.Cursor := crHourglass;
  try
    try
      WorksheetGrid.LoadFromSpreadsheetFile(UTF8ToSys(AFileName));

      // Update user interface
      Caption := Format('fpsGrid - %s (%s)', [
        AFilename,
        GetSpreadTechnicalName(WorksheetGrid.Workbook.FileFormatID)
      ]);

      // Collect the sheet names in the combobox for switching sheets.
      WorksheetGrid.GetSheets(SheetsCombo.Items);
      SheetsCombo.ItemIndex := 0;
    except
      on E:Exception do begin
        // Empty worksheet instead of the loaded one
        WorksheetGrid.NewWorkbook(26, 100);
        Caption := 'fpsGrid - no name';
        SheetsCombo.Items.Clear;
        // Grab the error message
        WorksheetGrid.Workbook.AddErrorMsg(E.Message);
      end;
    end;

  finally
    Screen.Cursor := crDefault;

    // Show a message in case of error(s)
    err := WorksheetGrid.Workbook.ErrorMsg;
    if err <> '' then
      MessageDlg(err, mtError, [mbOK], 0);
  end;
end;


initialization
  {$I mainform.lrs}

end.

