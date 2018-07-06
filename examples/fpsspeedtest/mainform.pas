unit mainform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, iniFiles,
  fpstypes, fpSpreadsheet;

type

  { TForm1 }

  TForm1 = class(TForm)
    Bevel1: TBevel;
    BtnSaveResults: TButton;
    BtnWrite: TButton;
    BtnRead: TButton;
    CgFormats: TCheckGroup;
    CgRowCount: TCheckGroup;
    CbVirtualModeOnly: TCheckBox;
    LblCancel: TLabel;
    Panel1: TPanel;
    Memo: TMemo;
    ParameterPanel: TPanel;
    RgColCount: TRadioGroup;
    RgContent: TRadioGroup;
    SaveDialog: TSaveDialog;
    StatusBar: TStatusBar;
    procedure BtnReadClick(Sender: TObject);
    procedure BtnSaveResultsClick(Sender: TObject);
    procedure BtnWriteClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
  private
    { private declarations }
    FDir: String;
    FEscape: Boolean;
    FCurFormat: TsSpreadsheetFormat;
    FErrCounter: Integer;
    FErrLog: TStringList;
    procedure BoldCheckgroup(AGroup: TCheckGroup);
    procedure BoldRadiogroup(AGroup: TRadioGroup);
    procedure EnableControls(AEnable: Boolean);
    function  GetRowCount(AIndex: Integer): Integer;
    procedure ReadCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
      const ADataCell: PCell);
    procedure WriteCellStringHandler(Sender: TsWorksheet;
      ARow, ACol: Cardinal; var AValue: Variant; var AStyleCell: PCell);
    procedure WriteCellNumberHandler(Sender: TsWorksheet;
      ARow, ACol: Cardinal; var AValue: Variant; var AStyleCell: PCell);
    procedure WriteCellStringAndNumberHandler(Sender: TsWorksheet;
      ARow, ACol: Cardinal; var AValue: Variant; var AStyleCell: PCell);
    procedure ReadFromIni;
    procedure WriteToIni;
    procedure RunReadTest(Idx: Integer; Log: String; Options: TsWorkbookOptions);
    procedure RunWriteTest(Idx: integer; Rows: integer; Log: string; Options: TsWorkbookOptions);
    procedure StatusMsg(const AMsg: String);
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

uses
  LclIntf, StrUtils, fpsUtils;

{$R *.lfm}

const
  fmtODS  = 0;
  fmtXLSX = 1;
  fmtXLS8 = 2;
  fmtXLS5 = 3;
  fmtXLS2 = 4;
  fmtCSV  = 5;

  rc10k = 0;
  rc20k = 1;
  rc30k = 2;
  rc40k = 3;
  rc50k = 4;
  rc60k = 5;
  rc100k = 6;

  CONTENT_PREFIX: array[0..2] of Char = ('S', 'N', 'M');
  CONTENT_TEXT: array[0..2] of string = ('strings only', 'numbers only', '50% strings and 50% numbers');

  FORMAT_EXT: array[0..5] of String = ('.ods', '.xlsx', '.xls', '_b5.xls', '_b2.xls', '.csv');
  SPREAD_FORMAT: array[0..5] of TsSpreadsheetFormat = (sfOpenDocument, sfOOXML, sfExcel8, sfExcel5, sfExcel2, sfCSV);

  COLCOUNT: array[0..2] of Integer = (1, 10, 100);

function LeftPad(AText: String; ALength: Integer): String;
begin
  Result := AText;
  while Length(Result) < ALength do Result := ' ' + Result;
end;


{ TForm1 }

procedure TForm1.ReadCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
  const ADataCell: PCell);
begin
  Unused(ACol, ADataCell);
  // nothing to do here. Just do a progress display
  if ARow mod 1000 = 0 then
    StatusMsg(Format('Virtual mode reading %s: Row %d...', [GetFileFormatName(FCurFormat), ARow]));
end;

procedure TForm1.WriteCellStringHandler(Sender: TsWorksheet; ARow, ACol: cardinal;
  var AValue: variant; var AStyleCell: PCell);
var
  S: string;
begin
  Unused(AStyleCell);
  S := 'Xy' + IntToStr(ARow) + 'x' + IntToStr(ACol);
  AValue := S;
  if ARow mod 1000 = 0 then
    StatusMsg(Format('Virtual mode writing %s: Row %d...', [GetFileFormatName(FCurFormat), ARow]));
end;

procedure TForm1.WriteCellNumberHandler(Sender: TsWorksheet; ARow, ACol: cardinal;
  var AValue: variant; var AStyleCell: PCell);
begin
  UnUsed(AStyleCell);
  AValue := ARow * 1E5 + ACol;
  if ARow mod 1000 = 0 then
    StatusMsg(Format('Virtual mode writing %s: Row %d...', [GetFileFormatName(FCurFormat), ARow]));
end;

procedure TForm1.WriteCellStringAndNumberHandler(Sender: TsWorksheet;
  ARow, ACol: cardinal; var AValue: variant; var AStyleCell: PCell);
begin
  if odd(ARow + ACol) then
    WriteCellStringHandler(Sender, ARow, ACol, AValue, AStyleCell)
  else
    WriteCellNumberHandler(Sender, ARow, ACol, AValue, AStyleCell);
end;

procedure TForm1.RunReadTest(Idx: Integer; Log: String;
  Options: TsWorkbookOptions);
var
  MyWorkbook: TsWorkbook;
  Tm: DWord;
  fName, s: String;
  k: Integer;
  F: File;
begin
  s := Trim(Log);     // needed for file name generation
  Log := Log + '         ';
  try
    if FEscape then begin
      Log := 'Test aborted';
      exit;
    end;

    for k := 0 to CgFormats.Items.Count-1 do begin
      if FEscape then begin
        Log := 'Test aborted';
        exit;
      end;

      if not CgFormats.Checked[k] then
        continue;

      FCurFormat := SPREAD_FORMAT[k];
      StatusMsg('Reading ' + GetFileFormatName(FCurFormat));

      if (SPREAD_FORMAT[k] = sfCSV) and (idx in [2, 4, 6, 8]) then
      begin
        Log := Log + '   -   ';              // No virtual mode for CSV
        continue;
      end;

      fName := FDir + CONTENT_PREFIX[RgContent.ItemIndex] + Copy(s, 1, Pos(' ', s)-1) + '_' + IntToStr(idx) + FORMAT_EXT[k];
      if not FileExists(fname) then
      begin
        inc(FErrCounter);
        FErrLog.Add(Format('[%d] = File not found.', [FErrCounter]));
        Log := Log + LeftPad(Format('[%d]',[FErrCounter]),5) + '  ';
        continue;
      end;
      AssignFile(F, fname);
      Reset(F);
      if FileSize(F) = 0 then
      begin
        inc(FErrCounter);
        FErrLog.Add(Format('[%d] = File size zero.', [FErrCounter]));
        Log := Log + LeftPad(Format('[%d]',[FErrCounter]),5) + '  ';
        continue;
      end;
      CloseFile(F);

      MyWorkbook := TsWorkbook.Create;
      try
        Application.ProcessMessages;
        MyWorkbook.Options := Options;
        if boVirtualMode in Options then
          MyWorkbook.OnReadCellData := @ReadCellDataHandler;
        Tm := GetTickCount;
        try
          MyWorkbook.ReadFromFile(fname, SPREAD_FORMAT[k]);
          Log := Log + format('%5.1f  ', [(GetTickCount - Tm) / 1000]);
        except
          on E:Exception do begin
            inc(FErrCounter);
            FErrLog.Add(Format('[%d] = %s', [FErrCounter, E.Message]));
            Log := Log + LeftPad(Format('[%d]',[FErrCounter]),5) + '  ';
          end;
        end;
      finally
        MyWorkbook.Free;
      end;
    end;

  finally
    Memo.Append(TrimRight(Log));
    StatusMsg('');
  end;
end;

procedure TForm1.RunWriteTest(Idx: integer; Rows: integer; Log: string;
  Options: TsWorkbookOptions);
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  ARow, ACol: cardinal;
  Tm: DWORD;
  fName, S: string;
  k: Integer;
  numCols: Integer;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    if FEscape then begin
      Log := 'Test aborted';
      exit;
    end;

    numCols := COLCOUNT[RgColCount.ItemIndex];

    MyWorksheet := MyWorkbook.AddWorksheet('Sheet1');
    MyWorkbook.Options := Options;

    Application.ProcessMessages;
    Tm := GetTickCount;

    try
      if boVirtualMode in Options then
      begin
        MyWorksheet.VirtualRowCount := Rows;
        MyWorksheet.VirtualColCount := numCols;
        case RgContent.ItemIndex of
          0: MyWorksheet.OnWriteCellData := @WriteCellStringHandler;
          1: MyWorksheet.OnWriteCellData := @WriteCellNumberHandler;
          2: MyWorksheet.OnWriteCellData := @WriteCellStringAndNumberHandler;
        end;
      end
      else
      begin
        for ARow := 0 to Rows - 1 do
        begin
          if ARow mod 1000 = 0 then begin
            StatusMsg(Format('Building row %d...', [ARow]));
            if FEscape then begin
              Log := 'Test aborted';
              exit;
            end;
          end;
          case RgContent.ItemIndex of
            0: for ACol := 0 to numCols-1 do begin
                 S := 'Xy' + IntToStr(ARow) + 'x' + IntToStr(ACol);
                 MyWorksheet.WriteText(ARow, ACol, S);
               end;
            1: for ACol := 0 to numCols-1 do
                 MyWorksheet.WriteNumber(ARow, ACol, 1E5*ARow + ACol);
            2: for ACol := 0 to numCols-1 do
                 if (odd(ARow) and odd(ACol)) or odd(ARow+ACol) then
                 begin
                   S := 'Xy' + IntToStr(ARow) + 'x' + IntToStr(ACol);
                   MyWorksheet.WriteText(ARow, ACol, S);
                 end else
                   MyWorksheet.WriteNumber(ARow, ACol, 1E5*ARow + ACol);
          end;
        end;
      end;
    except
      on E: Exception do begin
        inc(FErrCounter);
        FErrLog.Add(Format('[%d] = %s', [FErrCounter, E.Message]));
        Log := Log + LeftPad(Format('[%d]',[FErrCounter]),5) + '  ';
      end;
    end;

    fname :=  Trim(Log);
    fname := CONTENT_PREFIX[RgContent.ItemIndex] + copy(fname, 1, pos(' ', fname)-1);
    fname := FDir + fname + '_' + IntToStr(idx);

    if Idx in [2, 4, 6, 8]  then
      Log := Log +  '     -   '    // No build time in virtual mode
    else
      Log := Log + '  ' + Format('%5.1f  ', [(GetTickCount - Tm) / 1000]);

    for k := 0 to CgFormats.Items.Count-1 do
    begin
      if FEscape then begin
        Log := 'Test aborted';
        exit;
      end;

      if not CgFormats.Checked[k] then
        continue;

      FCurFormat := SPREAD_FORMAT[k];

      StatusMsg('Writing ' + GetFileFormatName(SPREAD_FORMAT[k]));
      try
        Application.ProcessMessages;
        Tm := GetTickCount;
        if (SPREAD_FORMAT[k] = sfCSV) and (Idx in [2, 4, 6, 8]) then
          Log := Log + '   -   '              // No virtual mode for CSV
        else begin
          MyWorkbook.WriteToFile(fname + FORMAT_EXT[k], SPREAD_FORMAT[k], true);
          Log := Log + Format('%5.1f  ', [(GetTickCount - Tm) / 1000]);
        end;
      except
        on E: Exception do
          Log := Log + ' xxxx  ';
      end;
    end;

  finally
    MyWorkbook.Free;
    Memo.Append(TrimRight(Log));
    StatusMsg('');
  end;

end;

procedure TForm1.StatusMsg(const AMsg: String);
begin
  Statusbar.SimpleText := AMsg;
  Statusbar.Refresh;
end;

function TForm1.GetRowCount(AIndex: Integer): Integer;
var
  s: String;
begin
  s := CgRowCount.Items[AIndex];
  Delete(s, pos('k', s), 99);
  Result := StrToInt(s) * 1000;
end;

procedure TForm1.BoldCheckgroup(AGroup: TCheckGroup);
var
  i: Integer;
begin
  for i:=0 to AGroup.ControlCount-1 do
    TCheckbox(AGroup.Controls[i]).ParentFont := false;
  AGroup.Font.Style := [fsBold];
end;

procedure TForm1.BoldRadiogroup(AGroup: TRadioGroup);
var
  i: Integer;
begin
  for i:=0 to AGroup.ControlCount-1 do
    TRadioButton(AGroup.Controls[i]).ParentFont := false;
  AGroup.Font.Style := [fsBold];
end;

procedure TForm1.BtnReadClick(Sender: TObject);
var
  i, len: Integer;
  s: String;
  rows: Integer;
  numCols: Integer;
begin
  WriteToIni;

  FEscape := false;
  EnableControls(false);
  FErrLog.Clear;

  numCols := COLCOUNT[RgColCount.ItemIndex];

  Memo.Append     ('Running: Reading TsWorkbook from various file formats');
  Memo.Append     ('         Worksheet contains ' + CONTENT_TEXT[RgContent.ItemIndex]);
  Memo.Append     ('         (Times in seconds)');
     //'-----------                        .ods  .xlsx  biff8  biff5  biff2');
     //'Rows x Cols  Options       Build  Write  Write  Write  Write  Write'
  s := '---------------------------------------  ';
  if CgFormats.Checked[fmtODS]  then s := s + ' .ods  ';
  if CgFormats.Checked[fmtXLSX] then s := s + '.xlsx  ';
  if CgFormats.Checked[fmtXLS8] then s := s + 'biff8  ';
  if CgFormats.Checked[fmtXLS5] then s := s + 'biff5  ';
  if CgFormats.Checked[fmtXLS2] then s := s + 'biff2  ';
  if CgFormats.Checked[fmtCSV]  then s := s + ' .csv';
  Memo.Append(TrimRight(s));
  s := 'Rows x Cols  Options                     ';
  if CgFormats.Checked[fmtODS]  then s := s + ' Read  ';
  if CgFormats.Checked[fmtXLSX] then s := s + ' Read  ';
  if CgFormats.Checked[fmtXLS8] then s := s + ' Read  ';
  if CgFormats.Checked[fmtXLS5] then s := s + ' Read  ';
  if CgFormats.Checked[fmtXLS2] then s := s + ' Read  ';
  if CgFormats.Checked[fmtCSV]  then s := s + ' Read';
  s := TrimRight(s);
  Memo.Append(s);
  len := Length(s);
  Memo.Append(DupeString('-', len));

  try
    for i:=0 to CgRowCount.Items.Count-1 do begin
      if FEscape then
        exit;

      if not CgRowCount.Checked[i] then
        continue;

      rows := GetRowCount(i);
      s := Format('%7.0nx%d', [1.0*rows, numCols]);
      if numCols < 10 then s := s + '  ' else if numCols < 100 then s := s + ' ';

      if CbVirtualModeOnly.Checked then begin
        RunReadTest(2, s + '  [boVM            ]', [boVirtualMode]);
        RunReadTest(4, s + '  [boVM, boBS      ]', [boVirtualMode, boBufStream]);
        RunReadTest(6, s + '  [boVM,       boFS]', [boVirtualMode, boFileStream]);
      end else begin
        RunReadTest(1, s + '  [                ]', []);
        RunReadTest(2, s + '  [boVM            ]', [boVirtualMode]);
        RunReadTest(3, s + '  [      boBS      ]', [boBufStream]);
        RunReadTest(4, s + '  [boVM, boBS      ]', [boVirtualMode, boBufStream]);
        RunReadTest(5, s + '  [            boFS]', [boFileStream]);
        RunReadTest(6, s + '  [boVM,       boFS]', [boVirtualMode, boFileStream]);
      end;

      Memo.Append(DupeString('-', len));
    end;
    if FErrLog.Text <> '' then
      Memo.Append(FErrLog.Text);
    Memo.Append('Ready');
  finally
    Memo.Append('');
    EnableControls(true);
  end;
end;

procedure TForm1.BtnSaveResultsClick(Sender: TObject);
begin
  if SaveDialog.Filename <> '' then
    SaveDialog.InitialDir := ExtractFileDir(SaveDialog.Filename);
  if SaveDialog.Execute then
    Memo.Lines.SaveToFile(SaveDialog.Filename);
end;

procedure TForm1.BtnWriteClick(Sender: TObject);
var
  Rows: integer;
  s: String;
  i, len: Integer;
  numCols: Integer;
begin
  WriteToIni;

  FEscape := false;
  EnableControls(false);
  FErrLog.Clear;
  numCols := COLCOUNT[RgColCount.ItemIndex];

  Memo.Append     ('Running: Building TsWorkbook and writing to different file formats');
  Memo.Append     ('         Worksheet contains ' + CONTENT_TEXT[RgContent.ItemIndex]);
  Memo.Append     ('         (Times in seconds)');
     //'-----------                        .ods  .xlsx  biff8  biff5  biff2');
     //'Rows x Cols  Options       Build  Write  Write  Write  Write  Write'
  s := '---------------------------------------  ';
  if CgFormats.Checked[fmtODS]  then s := s + ' .ods  ';
  if CgFormats.Checked[fmtXLSX] then s := s + '.xlsx  ';
  if CgFormats.Checked[fmtXLS8] then s := s + 'biff8  ';
  if CgFormats.Checked[fmtXLS5] then s := s + 'biff5  ';
  if CgFormats.Checked[fmtXLS2] then s := s + 'biff2  ';
  if CgFormats.Checked[fmtCSV]  then s := s + ' .csv';
  Memo.Append(TrimRight(s));
  s := 'Rows x Cols  Options              Build  ';
  if CgFormats.Checked[fmtODS]  then s := s + 'Write  ';
  if CgFormats.Checked[fmtXLSX] then s := s + 'Write  ';
  if CgFormats.Checked[fmtXLS8] then s := s + 'Write  ';
  if CgFormats.Checked[fmtXLS5] then s := s + 'Write  ';
  if CgFormats.Checked[fmtXLS2] then s := s + 'Write  ';
  if CgFormats.Checked[fmtCSV]  then s := s + 'Write';
  s := TrimRight(s);
  len := Length(s);
  Memo.Append(s);
  Memo.Append(DupeString('-', len));

  try
    for i:=0 to CgRowCount.Items.Count-1 do begin
      if FEscape then
        exit;

      if not CgRowCount.Checked[i] then
        continue;
      Rows := GetRowCount(i);
      s := Format('%7.0nx%d', [1.0*Rows, numCols]);
      if numCols < 10 then s := s + '  ' else if numCols < 100 then s := s + ' ';
      if CbVirtualModeOnly.Checked then begin
        RunWriteTest(2, Rows, s + '  [boVM            ]', [boVirtualMode]);
        RunWriteTest(4, Rows, s + '  [boVM, boBS      ]', [boVirtualMode, boBufStream]);
        RunWriteTest(6, Rows, s + '  [boVM,       boFS]', [boVirtualMode, boFileStream]);
      end else begin
        RunWriteTest(1, Rows, s + '  [                ]', []);
        RunWriteTest(2, Rows, s + '  [boVM            ]', [boVirtualMode]);
        RunWriteTest(3, Rows, s + '  [      boBS      ]', [boBufStream]);
        RunWriteTest(4, Rows, s + '  [boVM, boBS      ]', [boVirtualMode, boBufStream]);
        RunWriteTest(5, Rows, s + '  [            boFS]', [boFileStream]);
        RunWriteTest(6, Rows, s + '  [boVM,       boFS]', [boVirtualMode, boFileStream]);
      end;
      Memo.Append(DupeString('-', len));
    end;
    Memo.Append('Ready');
  finally
    Memo.Append('');
    EnableControls(true);
  end;

end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if CanClose then
    try
      WriteToIni;
    except
    end;
end;

procedure TForm1.EnableControls(AEnable: Boolean);
begin
  BtnWrite.Enabled := AEnable;
  BtnRead.Enabled := AEnable;
  BtnSaveResults.Enabled := AEnable;
  RgContent.Enabled := AEnable;
  CgFormats.Enabled := AEnable;
  CgRowCount.Enabled := AEnable;
  RgColCount.Enabled := AEnable;
  LblCancel.Visible := not AEnable;
  StatusMsg('');
  Application.ProcessMessages;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  //FDir := GetTempDir;
  FDir := ExtractFilePath(Application.ExeName) + 'data' + DirectorySeparator;
    // better than tempdir if you want to look at the files written...
  if not DirectoryExists(FDir) then CreateDir(FDir);

  CgFormats.Checked[fmtODS] := true;
  CgFormats.Checked[fmtXLSX] := true;
  CgFormats.Checked[fmtXLS8] := true;
  CgFormats.Checked[fmtXLS5] := true;
  CgFormats.Checked[fmtXLS2] := true;
  CgFormats.Checked[fmtCSV] := true;

  CgRowCount.Checked[rc10k] := true;
  CgRowCount.Checked[rc20k] := true;
  CgRowCount.Checked[rc30k] := true;
  CgRowCount.Checked[rc40k] := true;

  FErrLog := TStringList.Create;

  ReadFromIni;

  BoldRadiogroup(RgContent);
  BoldRadiogroup(RgColCount);
  BoldCheckGroup(CgFormats);
  BoldCheckGroup(CgRowCount);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FErrLog.Free;
end;

procedure TForm1.FormKeyPress(Sender: TObject; var Key: char);
begin
  if Key = #27 then begin
    StatusMsg('ESC pressed...');
    FEscape := true;
  end;
end;

procedure TForm1.ReadFromIni;
var
  ini: TCustomIniFile;
  n: Byte;
begin
  ini := TMemIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try
    CbVirtualModeOnly.Checked := ini.ReadBool('Parameters', 'VirtualModeOnly', CbVirtualModeOnly.Checked);
    RgContent.ItemIndex := ini.ReadInteger('Parameters', 'Content', RgContent.ItemIndex);
    RgColCount.ItemIndex := Ini.ReadInteger('Parameters', 'ColCount', RgColCount.ItemIndex);

    n := Ini.ReadInteger('Parameters', 'Formats', $1F);
    CgFormats.Checked[fmtODS]  := n and $01 <> 0;
    CgFormats.Checked[fmtXLSX] := n and $02 <> 0;
    CgFormats.Checked[fmtXLS8] := n and $04 <> 0;
    CgFormats.Checked[fmtXLS5] := n and $08 <> 0;
    CgFormats.Checked[fmtXLS2] := n and $10 <> 0;
    CgFormats.Checked[fmtCSV]  := n and $20 <> 0;

    n := Ini.ReadInteger('Parameters', 'RowCount', $0F);
    CgRowCount.Checked[rc10k] := n and $01 <> 0;
    CgRowCount.Checked[rc20k] := n and $02 <> 0;
    CgRowCount.Checked[rc30k] := n and $04 <> 0;
    CgRowCount.Checked[rc40k] := n and $08 <> 0;
    CgRowCount.Checked[rc50k] := n and $10 <> 0;
    CgRowCount.Checked[rc60k] := n and $20 <> 0;
    CgRowCount.Checked[rc100k]:= n and $40 <> 0;

  finally
    ini.Free;
  end;
end;

procedure TForm1.WriteToIni;
var
  ini: TMemIniFile;
  n: Byte;
begin
  ini := TMemIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try
    ini.WriteBool   ('Parameters', 'VirtualModeOnly', CbVirtualModeOnly.Checked);
    ini.WriteInteger('Parameters', 'Content', RgContent.ItemIndex);
    ini.WriteInteger('Parameters', 'ColCount', RgColCount.ItemIndex);

    n := 0;
    if CgFormats.Checked[fmtODS]  then n := n or $1;
    if CgFormats.Checked[fmtXLSX] then n := n or $2;
    if CgFormats.Checked[fmtXLS8] then n := n or $4;
    if CgFormats.Checked[fmtXLS5] then n := n or $8;
    if CgFormats.Checked[fmtXLS2] then n := n or $10;
    if CgFormats.Checked[fmtCSV]  then n := n or $20;
    ini.WriteInteger('Parameters', 'Formats', n);

    n := 0;
    if CgRowCount.Checked[rc10k]  then n := n or $01;
    if CgRowCount.Checked[rc20k]  then n := n or $02;
    if CgRowCount.Checked[rc30k]  then n := n or $04;
    if CgRowCount.Checked[rc40k]  then n := n or $08;
    if CgRowCount.Checked[rc50k]  then n := n or $10;
    if CgRowCount.Checked[rc60k]  then n := n or $20;
    if CgRowCount.Checked[rc100k] then n := n or $40;
    ini.WriteInteger('Parameters', 'RowCount', n);

  finally
    ini.UpdateFile;
    ini.Free;
  end;
end;

end.
