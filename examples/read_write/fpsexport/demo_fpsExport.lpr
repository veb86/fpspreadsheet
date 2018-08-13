program demo_fpsExport;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Classes
  { you can add units after this },
  SysUtils, db, bufdataset,
  fpspreadsheet, fpsexport;

type
  TExportTestData=record
    id: integer;
    Name: string;
    DOB: TDateTime;
  end;

var
  ExportTestData: array[0..5] of TExportTestData;
  FDataset: TBufDataset;

procedure InitExportData;
begin
  with ExportTestData[0] do
  begin
    id:=1;
    name:='Elvis Wesley';
    dob := encodedate(1912,12,31);
  end;

  with ExportTestData[1] do
  begin
    id:=2;
    name:='Kingsley Dill';
    dob:=encodedate(1918,11,11);
  end;

  with ExportTestData[2] do
  begin
    id:=3;
    name:='Joe Snort';
    dob:=encodedate(1988,8,4);
  end;

  with ExportTestData[3] do
  begin
    id:=4;
    name:='Hagen Dit';
    dob:=encodedate(1944,2,24);
  end;

  with ExportTestData[4] do
  begin
    id:=5;
    name:='Kingsley Snort';
    dob:=encodedate(1928,11,11);
  end;

  with ExportTestData[5] do
  begin
    id:=6;
    name:='';
    dob:=encodedate(2112,4,12);
  end;
end;

procedure CreateDB;
var
  i:integer;
begin
  FDataset:=TBufDataset.Create(nil);
  with FDataset.FieldDefs do
  begin
    Add('id',ftAutoinc);
    Add('name',ftString,40);
    Add('dob',ftDateTime);
  end;
  FDataset.CreateDataset;

  for i:=low(ExportTestData) to high(ExportTestData) do
  begin
    FDataset.Append;
    //autoinc field should be filled by bufdataset
    FDataSet.Fields.FieldByName('name').AsString:=ExportTestData[i].Name;
    FDataSet.Fields.FieldByName('dob').AsDateTime:=ExportTestData[i].dob;
    FDataSet.Post;
  end;
end;

procedure Cleanup;
begin
  FDataset.Free;
end;

procedure ExportDatabaseToSpreadsheet(AFilename: String);
var
  Exp: TFPSExport;
  ExpSettings: TFPSExportFormatSettings;
begin
  FDataset.First;
  Exp := TFPSExport.Create(nil);
  ExpSettings := TFPSExportFormatSettings.Create(true);
  try
    ExpSettings.ExportFormat := efXLS;
    ExpSettings.HeaderRow := true;
    Exp.FormatSettings := ExpSettings;
    Exp.Dataset := FDataset;
    Exp.FileName := AFileName;
    Exp.Execute;
  finally
    Exp.Free;
    ExpSettings.Free;
  end;
end;

begin
  WriteLn('Preparing test data...');
  InitExportData;
  CreateDB;
  WriteLn('Exporting...');
  ExportDatabaseToSpreadsheet('temp.xls');
  WriteLn('Done.');
  Cleanup;
  {$IFDEF WINDOWS}
  WriteLn;
  WriteLn('Press RETURN to quit.');
  ReadLn;
  {$ENDIF}
end.

