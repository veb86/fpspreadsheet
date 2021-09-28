unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, Forms, Controls, Graphics, Dialogs, DBGrids, StdCtrls,
  ExtCtrls, fpsDataset;

type

  { TMainForm }

  TMainForm = class(TForm)
    btnViewSpreadsheet: TButton;
    DataSource: TDataSource;
    Dataset: TsWorksheetDataset;
    DBGrid: TDBGrid;
    Panel1: TPanel;
    procedure btnViewSpreadsheetClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private

  public

  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

uses
  ViewerForm;

const
  FILE_NAME = 'TestData.ods';

{ TMainForm }

procedure TMainForm.FormCreate(Sender: TObject);
begin
  // For demonstration purposes we want to create a new file whenever the demo starts.
  if FileExists(FILE_NAME) then
    DeleteFile(FILE_NAME);

  // Set the name of the data file
  Dataset.FileName := FILE_NAME;

  // Set the name of the worksheet
  Dataset.SheetName := 'Test-Sheet';

  // Define fields
  Dataset.AutoFieldDefs := false;
  Dataset.AddFieldDef('FloatColumn', ftFloat);
  Dataset.AddFieldDef('StringColumn', ftString, 20);
  Dataset.AddFieldDef('DateColumn', ftDate);
  Dataset.AddFieldDef('BooleanColumn', ftBoolean, 0, 4);  // 4 --> column 4 in worksheet --> skip column 3
  Dataset.CreateTable;

  // Open the table
  Dataset.Open;
  (Dataset.FieldByName('FloatColumn') as TNumericField).DisplayFormat := '0.000';

  // Add some arbitrary data
  Dataset.Append;
  Dataset.FieldByName('FloatColumn').AsFloat := 3.1415;
  Dataset.FieldByName('StringColumn').AsString := 'abc';
  Dataset.FieldByName('DateColumn').AsDateTime := EncodeDate(2000, 1, 1);
  Dataset.FieldByName('BooleanColumn').AsBoolean := true;
  Dataset.Post;

  Dataset.Append;
  Dataset.FieldByName('FloatColumn').AsFloat := 2*3.1415;
  Dataset.FieldByName('StringColumn').AsString := 'Lorem ipsum';
  Dataset.FieldByName('DateColumn').AsDateTime := Date();
  Dataset.FieldByName('BooleanColumn').AsBoolean := false;
  Dataset.Post;

  Dataset.Append;
  Dataset.FieldByName('FloatColumn').AsFloat := 3*3.1415;
  Dataset.FieldByName('StringColumn').AsString := 'Hello World';
  Dataset.FieldByName('DateColumn').AsDateTime := Date() + 1;
  Dataset.FieldByName('BooleanColumn').AsInteger := 3;  // anything <> 0 is "true"
  Dataset.Post;
end;

procedure TMainForm.btnViewSpreadsheetClick(Sender: TObject);
begin
  Dataset.Flush;

  if SpreadsheetViewerForm = nil then
    SpreadsheetViewerForm := TSpreadsheetViewerForm.Create(Application);

  SpreadsheetViewerForm.LoadFile(Dataset.FileName, Dataset.SheetName);
  SpreadsheetViewerForm.Show;
end;

end.

