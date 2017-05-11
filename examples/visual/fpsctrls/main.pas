unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, ActnList, Menus, StdActns, Buttons,
  fpstypes, fpspreadsheet, fpspreadsheetctrls, fpspreadsheetgrid, fpsActions,
  fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    ActionList1: TActionList;
    FileExit1: TFileExit;
    FileOpen1: TFileOpen;
    FileSaveAs1: TFileSaveAs;
    ImageList1: TImageList;
    MainMenu1: TMainMenu;
    MnuFile: TMenuItem;
    MnuFileOpen: TMenuItem;
    MnuFileSaveAs: TMenuItem;
    MnuFileSeparator: TMenuItem;
    MnuFileExit: TMenuItem;
    MnuFormat: TMenuItem;
    MnuFormatBold: TMenuItem;
    MnuFormatItalic: TMenuItem;
    MnuFormatUnderline: TMenuItem;
    sWorkbookSource1: TsWorkbookSource;
    sWorkbookTabControl1: TsWorkbookTabControl;
    sWorksheetGrid1: TsWorksheetGrid;
    sCellEdit1: TsCellEdit;
    sCellIndicator1: TsCellIndicator;
    sFontStyleAction1: TsFontStyleAction;
    sFontStyleAction2: TsFontStyleAction;
    sFontStyleAction3: TsFontStyleAction;
    Panel2: TPanel;
    FontColorCombobox: TsCellCombobox;
    FontNameCombo: TsCellCombobox;
    FontSizeCombo: TsCellCombobox;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    ToolBar3: TToolBar;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    procedure FileOpen1Accept(Sender: TObject);
    procedure FileSaveAs1Accept(Sender: TObject);
  private

  protected

  public

  end;

var
  Form1: TForm1;


implementation

{$R *.lfm}

uses
  LCLIntf,
  fpsUtils, fpsCSV;


{ TForm1 }

{ Loads the spreadsheet file selected by the FileOpen standard action }
procedure TForm1.FileOpen1Accept(Sender: TObject);
begin
  sWorkbookSource1.AutodetectFormat := false;
  case FileOpen1.Dialog.FilterIndex of
    1: sWorkbookSource1.AutoDetectFormat := true;         // All spreadsheet files
    2: sWorkbookSource1.AutoDetectFormat := true;         // All Excel files
    3: sWorkbookSource1.FileFormat := sfOOXML;            // Excel 2007+
    4: sWorkbookSource1.FileFormat := sfExcel8;           // Excel 97-2003
    5: sWorkbookSource1.FileFormat := sfExcel5;           // Excel 5.0
    6: sWorkbookSource1.FileFormat := sfExcel2;           // Excel 2.1
    7: sWorkbookSource1.FileFormat := sfOpenDocument;     // Open/LibreOffice
    8: sWorkbookSource1.FileFormat := sfCSV;              // Text files
  end;
  sWorkbookSource1.FileName :=FileOpen1.Dialog.FileName;  // This loads the file
end;

{ Saves the spreadsheet to the file selected by the FileSaveAs1 action }
procedure TForm1.FileSaveAs1Accept(Sender: TObject);
var
  fmt: TsSpreadsheetFormat;
begin
  Screen.Cursor := crHourglass;
  try
    case FileSaveAs1.Dialog.FilterIndex of
      1: fmt := sfOOXML;                // Note: Indexes are 1-based here!
      2: fmt := sfExcel8;
      3: fmt := sfExcel5;
      4: fmt := sfExcel2;
      5: fmt := sfOpenDocument;
      6: fmt := sfCSV;
      7: fmt := sfWikiTable_WikiMedia;
    end;
    sWorkbookSource1.SaveToSpreadsheetFile(FileSaveAs1.Dialog.FileName, fmt);
  finally
    Screen.Cursor := crDefault;
  end;
end;

end.

