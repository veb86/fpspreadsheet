unit mainfrm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, StdCtrls, Grids,
  fpstypes, fpspreadsheet, fpspreadsheetgrid, {%H-}fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    Bevel1: TBevel;
    BtnNew: TButton;
    BtnLoad: TButton;
    BtnSave: TButton;
    BtnEnterText: TButton;
    ButtonPanel: TPanel;
    CbAutoExpandOnData: TCheckBox;
    CbAutoExpandOnNavigation: TCheckBox;
    EdCellValue: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    SaveDialog: TSaveDialog;
    TabControl: TTabControl;
    procedure BtnEnterTextClick(Sender: TObject);
    procedure BtnLoadClick(Sender: TObject);
    procedure BtnNewClick(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure CbAutoExpandOnDataChange(Sender: TObject);
    procedure CbAutoExpandOnNavigationChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
  private
    { private declarations }
    Grid: TsWorksheetGrid;
    procedure LoadFile(const AFileName: String);
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

uses
  fpsUtils, fpsReaderWriter;


{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
const
  THICK_BORDER: TsCellBorderStyle = (LineStyle: lsThick; Color: clNavy);
  MEDIUM_BORDER: TsCellBorderSTyle = (LineStyle: lsMedium; Color: clRed);
  DOTTED_BORDER: TsCellBorderSTyle = (LineStyle: lsDotted; Color: clRed);
begin
  Grid := TsWorksheetGrid.Create(self);

  // Put the grid into the TabControl and align it to fill the tabcontrol.
  Grid.Parent := TabControl;
  Grid.Align := alClient;

  // Useful options and properties
  Grid.Options := Grid.Options + [goColSizing, goRowSizing,
//    goAutoAddRows,
//    goAutoAddRowsSkipContentCheck,
    goFixedColSizing,    // useful if the spreadsheet contains frozen columns
    goEditing,           // needed for modifying cell content
    goThumbTracking,     // see the grid scroll while you drag the scrollbar
    goHeaderHotTracking, // hot-tracking of header cells
    goHeaderPushedLook,  // click at header cells --> pushed look
    goDblClickAutoSize,  // optimum col width/row height after dbl click at header border
    goCellHints          // show cell hints (needed for cell comments)
  ];
  Grid.AutoAdvance := aaDown;       // on ENTER, move active cell down
  Grid.MouseWheelOption := mwGrid;  // mouse wheel scrolls the grid, not the active cell
  Grid.TextOverflow := true;        // too long text extends into neighbor cells
  Grid.AutoCalc := true;            // automatically calculate formulas
  Grid.ShowHint := true;            // needed to show cell comments
  Grid.RowCount := 10;              // Prepare 10 columns (incl fixed header)
  Grid.ColCount := 8;               // and 8 rows (incl fixed header) - but grid expands automatically

  CbAutoExpandOnData.Checked := aeData in Grid.AutoExpand;
  CbAutoExpandOnNavigation.Checked := aeNavigation in Grid.AutoExpand;

  // Add some cells and formats
  Grid.ColWidths[1] := 180;
  Grid.ColWidths[2] := 100;

  Grid.Cells[1,1] := 'This is a demo';
  Grid.MergeCells(1,1, 2,1);
  Grid.HorAlignment[1,1] := haCenter;
  Grid.CellBorders[1,1, 2,1] := [cbSouth];
  Grid.CellBorderStyles[1,1, 2,1, cbSouth] := THICK_BORDER;
  Grid.BackgroundColors[1,1, 2,1] := RGBToColor(232, 242, 255);
  Grid.CellFontColor[1,1] := clNavy;
  Grid.CellFontStyle[1,1] := [fssBold];

  Grid.Cells[1,2] := 'Number:';
  Grid.HorAlignment[1,2] := haRight;
  Grid.CellFontStyle[1,2] := [fssItalic];
  Grid.CellFontColor[1,2] := clNavy;
  Grid.Cells[2,2] := 1.234;

  Grid.Cells[1,3] := 'Date:';
  Grid.HorAlignment[1,3] := haRight;
  Grid.CellFontStyle[1,3] := [fssItalic];
  Grid.CellFontColor[1,3] := clNavy;
  Grid.NumberFormat[2,3] := 'mmm dd, yyyy';
  Grid.Cells[2,3] := date;

  Grid.Cells[1,4] := 'Time:';
  Grid.HorAlignment[1,4] := haRight;
  Grid.CellFontStyle[1,4] := [fssItalic];
  Grid.CellFontColor[1,4] := clNavy;
  Grid.NumberFormat[2,4] := 'hh:nn';
  Grid.Cells[2,4] := now();

  Grid.Cells[1,5] := 'Rich text:';
  Grid.HorAlignment[1,5] := haRight;
  Grid.CellFontStyle[1,5] := [fssItalic];
  Grid.CellFontColor[1,5] := clNavy;
  Grid.Cells[2,5] := '100 cm<sup>2</sup>';

  Grid.Cells[1,6] := 'Formula:';
  Grid.HorAlignment[1,6] := haRight;
  Grid.CellFontStyle[1,6] := [fssItalic];
  Grid.CellFontColor[1,6] := clNavy;
  Grid.Cells[2,6] := '=B2^2*PI()';
  Grid.CellComment[2,6] := 'Area of the circle with radius given in cell B2';
  Grid.NumberFormat[2,6] := '0.000';

  ActiveControl := Grid;
end;

procedure TForm1.BtnLoadClick(Sender: TObject);
begin
  if OpenDialog.FileName <> '' then
  begin
    OpenDialog.InitialDir := ExtractFileDir(OpenDialog.FileName);
    OpenDialog.FileName := ChangeFileExt(ExtractFileName(OpenDialog.FileName), '');
  end;
  if OpenDialog.Execute then
  begin
    LoadFile(OpenDialog.FileName);
  end;
end;

procedure TForm1.BtnEnterTextClick(Sender: TObject);
begin
  Grid.Worksheet.WriteText(109, 27, EdCellValue.Text);
  Grid.Worksheet.SelectCell(109, 27);
end;

procedure TForm1.BtnNewClick(Sender: TObject);
begin
  TabControl.Tabs.Clear;
  TabControl.Tabs.Add('Sheet1');
  Grid.NewWorkbook(26, 100);
end;

// Saves sheet in grid to file, overwriting existing file
procedure TForm1.BtnSaveClick(Sender: TObject);
var
  err: String;
  fn: String;
begin
  if Grid.Workbook = nil then
    exit;

  if Grid.Workbook.Filename <>'' then
  begin
    fn := AnsiToUtf8(Grid.Workbook.Filename);
    SaveDialog.InitialDir := ExtractFileDir(fn);
    SaveDialog.FileName := ChangeFileExt(ExtractFileName(fn), '');
  end;

  if SaveDialog.Execute then
  begin
    Screen.Cursor := crHourglass;
    try
      Grid.SaveToSpreadsheetFile(UTF8ToAnsi(SaveDialog.FileName));
    finally
      Screen.Cursor := crDefault;
      // Show a message in case of error(s)
      err := Grid.Workbook.ErrorMsg;
      if err <> '' then
        MessageDlg(err, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TForm1.CbAutoExpandOnDataChange(Sender: TObject);
begin
  if CbAutoExpandOnData.Checked then
    Grid.AutoExpand := Grid.AutoExpand + [aeData] else
    Grid.AutoExpand := Grid.AutoExpand - [aeData];
end;

procedure TForm1.CbAutoExpandOnNavigationChange(Sender: TObject);
begin
  if CbAutoExpandOnNavigation.Checked then
    Grid.AutoExpand := Grid.AutoExpand + [aeNavigation] else
    Grid.AutoExpand := Grid.AutoExpand - [aeNavigation];
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
      // Load file into workbook and grid
      Grid.LoadFromSpreadsheetFile(UTF8ToAnsi(AFileName));

      // Update user interface
      Caption := Format('fpsGrid - %s (%s)', [
        AFilename,
        GetSpreadTechnicalName(Grid.Workbook.FileFormatID)
      ]);

      // Collect the sheet names in the Tabs of the TabControl for switching sheets.
      Grid.GetSheets(TabControl.Tabs);
      TabControl.TabIndex := 0;
    except
      on E:Exception do begin
        // Empty worksheet instead of the loaded one
        Grid.NewWorkbook(26, 100);
        Caption := 'fpsGrid - no name';
        TabControl.Tabs.Clear;
        // Grab the error message, it will be displayed below
        Grid.Workbook.AddErrorMsg(E.Message);
      end;
    end;

  finally
    Screen.Cursor := crDefault;

    // Show a message in case of error(s)
    err := Grid.Workbook.ErrorMsg;
    if err <> '' then
      MessageDlg(err, mtError, [mbOK], 0);
  end;
end;

procedure TForm1.TabControlChange(Sender: TObject);
begin
  Grid.SelectSheetByIndex(TabControl.TabIndex);
end;


end.

