{ fpsReaderWriter }

{@@ ----------------------------------------------------------------------------
  Unit fpsReaderWriter implements basic reading/writing support
  for fpspreadsheet, as well as registration of the file formats supported.

  AUTHORS: Felipe Monteiro de Carvalho, Reinier Olislagers, Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.

  USAGE:   Each unit implementing a new spreadsheet format must register the
           reader/writer and some specific data by calling "RegisterSpreadFormat".
-------------------------------------------------------------------------------}

unit fpsReaderWriter;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

uses
  Classes, Sysutils, AVL_Tree,
  fpsTypes, fpsClasses, fpSpreadsheet;

type

  { TsBasicSpreadReaderWriter }
  TsBasicSpreadReaderWriter = class
  protected
    {@@ Instance of the workbook which is currently being read or written. }
    FWorkbook: TsWorkbook;
    {@@ Instance of the worksheet which is currently being read or written. }
    FWorksheet: TsWorksheet;
    {@@ Limitations for the specific data file format }
    FLimitations: TsSpreadsheetFormatLimitations;
  public
    constructor Create(AWorkbook: TsWorkbook); virtual;  // to allow descendents to override it
    function Limitations: TsSpreadsheetFormatLimitations;
    {@@ Instance of the workbook which is currently being read/written. }
    property Workbook: TsWorkbook read FWorkbook;
  end;

  { TsBasicSpreadReader }
  TsBasicSpreadReader = class(TsBasicSpreadReaderWriter)
  public
    { General writing methods }
    procedure ReadFromFile(AFileName: string; AParams: TsStreamParams = []); virtual; abstract;
    procedure ReadFromStream(AStream: TStream; AParams: TsStreamParams = []); virtual; abstract;
    procedure ReadFromStrings(AStrings: TStrings; AParams: TsStreamParams = []); virtual; abstract;
  end;

  { TsBasicSpreadWriter }
  TsBasicSpreadWriter = class(TsBasicSpreadReaderWriter)
  public
    { Helpers }
    procedure CheckLimitations; virtual;
    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False; AParams: TsStreamParams = []); virtual; abstract;
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); virtual; abstract;
    procedure WriteToStrings(AStrings: TStrings; AParams: TsStreamParams = []); virtual; abstract;
  end;

  {@@ TsSpreadReader class reference type }
  TsSpreadReaderClass = class of TsBasicSpreadReader;

  {@@ TsSpreadWriter class reference type }
  TsSpreadWriterClass = class of TsBasicSpreadWriter;

  {@@
    Custom reader of spreadsheet files. "Custom" means that it provides only
    the basic functionality. The main implementation is done in derived classes
    for each individual file format.
  }
  TsCustomSpreadReader = class(TsBasicSpreadReader)
  protected
    {@@ list of format records collected from the file }
    FCellFormatList: TsCellFormatList;
    {@@ List of fonts collected from the file }
    FFontList: TFPList;
    {@@ Temporary cell for virtual mode}
    FVirtualCell: TCell;
    {@@ Stores if the reader is in virtual mode }
    FIsVirtualMode: Boolean;
    {@@ List of number formats  }
    FNumFormatList: TStringList;

    { Helper methods }
    procedure AddBuiltinNumFormats; virtual;
    {@@ Removes column records if all of them have the same column width }
    procedure FixCols(AWorksheet: TsWorksheet);
    {@@ Removes row records if all of them have the same row height }
    procedure FixRows(AWorksheet: TsWorksheet);

    { Record reading methods }
    {@@ Abstract method for reading a blank cell. Must be overridden by descendent classes. }
    procedure ReadBlank(AStream: TStream); virtual; abstract;
    {@@ Abstract method for reading a BOOLEAN cell. Must be overridden by descendent classes. }
    procedure ReadBool(AStream: TSTream); virtual; abstract;
    {@@ Abstract method for reading a formula cell. Must be overridden by descendent classes. }
    procedure ReadFormula(AStream: TStream); virtual; abstract;
    {@@ Abstract method for reading a text cell. Must be overridden by descendent classes. }
    procedure ReadLabel(AStream: TStream); virtual; abstract;
    {@@ Abstract method for reading a number cell. Must be overridden by descendent classes. }
    procedure ReadNumber(AStream: TStream); virtual; abstract;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;

    { General writing methods }
    procedure ReadFromFile(AFileName: string; AParams: TsStreamParams = []); override;
    procedure ReadFromStream(AStream: TStream; AParams: TsStreamParams = []); override;
    procedure ReadFromStrings(AStrings: TStrings; AParams: TsStreamParams = []); override;

    {@@ List of number formats found in the workbook. }
    property NumFormatList: TStringList read FNumFormatList;
  end;


  {@@ Callback function when iterating cells while accessing a stream }
  TCellsCallback = procedure (ACell: PCell; AStream: TStream) of object;

  {@@ Callback function when iterating comments while accessing a stream }
  TCommentsCallback = procedure (AComment: PsComment; ACommentIndex: Integer;
    AStream: TStream) of object;

  {@@ Callback function when iterating hyperlinks while accessing a stream }
  THyperlinksCallback = procedure (AHyperlink: PsHyperlink;
    AStream: TStream) of object;

  {@@ Custom writer of spreadsheet files. "Custom" means that it provides only
    the basic functionality. The main implementation is done in derived classes
    for each individual file format. }
  TsCustomSpreadWriter = class(TsBasicSpreadWriter)
  protected
    {@@ List of number formats found in the file }
    FNumFormatList: TStringList;

    procedure AddBuiltinNumFormats; virtual;
    function  FindNumFormatInList(ANumFormatStr: String): Integer;
//    function  FixColor(AColor: TsColor): TsColor; virtual;
    procedure FixFormat(ACell: PCell); virtual;
    procedure GetSheetDimensions(AWorksheet: TsWorksheet;
      out AFirstRow, ALastRow, AFirstCol, ALastCol: Cardinal); virtual;
    procedure ListAllNumFormats; virtual;

    { Helpers for writing }
    procedure WriteCellToStream(AStream: TStream; ACell: PCell); virtual;
    procedure WriteCellsToStream(AStream: TStream; ACells: TsCells);

    { Record writing methods }
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); virtual; abstract;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); virtual; abstract;
    procedure WriteComment(AStream: TStream; ACell: PCell); virtual;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); virtual; abstract;
    procedure WriteError(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TsErrorValue; ACell: PCell); virtual; abstract;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); virtual;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); virtual; abstract;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); virtual; abstract;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;

    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False; AParams: TsStreamParams = []); override;
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); override;
    procedure WriteToStrings(AStrings: TStrings; AParams: TsStreamParams = []); override;

    {@@ List of number formats found in the workbook. }
    property NumFormatList: TStringList read FNumFormatList;
  end;


type
  TsSpreadFileAccess = (faRead, faWrite);

function RegisterSpreadFormat(
  AFormat: TsSpreadsheetFormat;
  AReaderClass: TsSpreadReaderClass;
  AWriterClass: TsSpreadWriterClass;
  AFormatName, ATechnicalName: String;
  const AFileExtensions: array of String): TsSpreadFormatID;

function GetFileFormatFilter(AListSeparator, AExtSeparator: Char;
  AFileAccess: TsSpreadFileAccess; const APriorityFormats: array of TsSpreadFormatID;
  AllSpreadFormats: Boolean = false; AllExcelFormats: Boolean = false): String;

function GetSpreadFormats(AFileAccess: TsSpreadFileAccess;
  const APriorityFormats: array of TsSpreadFormatID): TsSpreadFormatIDArray;
function GetSpreadFormatsFromFileName(AFileAccess: TsSpreadFileAccess; AFileName: TFileName;
  APriorityFormat: TsSpreadFormatID = sfidUnknown): TsSpreadFormatIDArray;

function GetSpreadFormatExt(AFormatID: TsSpreadFormatID): String;
function GetSpreadFormatName(AFormatID: TsSpreadFormatID): String;
function GetSpreadTechnicalName(AFormatID: TsSpreadFormatID): String;

function GetSpreadReaderClass(AFormatID: TsSpreadFormatID): TsSpreadReaderClass;
function GetSpreadWriterClass(AFormatID: TsSpreadFormatID): TsSpreadWriterClass;


implementation

uses
  Math, LazUTF8,
  fpsStrings, fpsUtils, fpsNumFormat, fpsStreams, fpsRegFileFormats;


{------------------------------------------------------------------------------}
{                          TsBasicSpreadReaderWriter                           }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the reader/writer. Has the workbook to be read/written as a
  parameter to apply the localization information found in its FormatSettings.

  @param AWorkbook  Workbook into which the file is being read or from with the
                    file is written. This parameter is passed from the workbook
                    which creates the reader/writer.
-------------------------------------------------------------------------------}
constructor TsBasicSpreadReaderWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  { A good starting point valid for many formats ... }
  FLimitations.MaxColCount := 256;
  FLimitations.MaxRowCount := 65536;
  FLimitations.MaxPaletteSize := MaxInt;
  FLimitations.MaxSheetnameLength := MaxInt;
end;

{@@ ----------------------------------------------------------------------------
  Returns a record containing limitations of the specific file format of the
  writer.
-------------------------------------------------------------------------------}
function TsBasicSpreadReaderWriter.Limitations: TsSpreadsheetFormatLimitations;
begin
  Result := FLimitations;
end;


{------------------------------------------------------------------------------}
{                             TsBasicSpreadWriter                              }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Checks limitations of the writer, e.g max row/column count
-------------------------------------------------------------------------------}
procedure TsBasicSpreadWriter.CheckLimitations;
var
  lastCol, lastRow: Cardinal;
  i: Integer;
  sheet: TsWorksheet;
begin
  Workbook.GetLastRowColIndex(lastRow, lastCol);

  // Check row count
  if lastRow >= FLimitations.MaxRowCount then
    Workbook.AddErrorMsg(rsMaxRowsExceeded, [lastRow+1, FLimitations.MaxRowCount]);

  // Check column count
  if lastCol >= FLimitations.MaxColCount then
    Workbook.AddErrorMsg(rsMaxColsExceeded, [lastCol+1, FLimitations.MaxColCount]);

  // Check worksheet names
  for i:=0 to Workbook.GetWorksheetCount-1 do
  begin
    sheet := Workbook.GetWorksheetByIndex(i);
    if UTF8Length(sheet.Name) > FLimitations.MaxSheetNameLength then
      // Worksheet name is too long.
      // We abort saving here because it is not safe to chop the sheet name
      // to its allowed length - it may be used as a reference in formulas.
      raise Exception.CreateFmt(rsWriteError_WorksheetNameTooLong,
        [sheet.Name, FLimitations.MaxSheetNameLength]);
  end;
end;


{------------------------------------------------------------------------------}
{                              TsCustomSpreadReader                            }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the reader. Has the workbook to be read as a
  parameter to apply the localization information found in its FormatSettings.
  Creates an internal instance of the number format list according to the
  file format being read/written.

  @param AWorkbook  Workbook into which the file is being read.
                    This parameter is passed from the workbook which creates
                    the reader.
-------------------------------------------------------------------------------}
constructor TsCustomSpreadReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  // Font list
  FFontList := TFPList.Create;
  // Number formats
  FNumFormatList := TStringList.Create;
  AddBuiltinNumFormats;
  // Virtual mode
  FIsVirtualMode := (boVirtualMode in FWorkbook.Options) and
    Assigned(FWorkbook.OnReadCellData);
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the reader. Destroys the internal number format list and the
  error log list.
-------------------------------------------------------------------------------}
destructor TsCustomSpreadReader.Destroy;
var
  j: Integer;
begin
  for j:=FFontList.Count-1 downto 0 do
    if FFontList[j] <> nil then TObject(FFontList[j]).Free;   // font #4 can add a nil!
  FreeAndNil(FFontList);

  FreeAndNil(FNumFormatList);
  FreeAndNil(FCellFormatList);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in number formats to the internal NumFormatList.

  Must be overridden by descendants because they know about the details of
  the file format.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadReader.AddBuiltinNumFormats;
begin
  // to be overridden by descendants
end;

{@@ ----------------------------------------------------------------------------
  Deletes unnecessary column records as they are written by some
  Office applications when they convert a file to another format.

  @param   AWorksheet   The columns in this worksheet are processed.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadReader.FixCols(AWorkSheet: TsWorksheet);
const
  EPS = 1E-3;
var
  c: LongInt;
  w: Single;
  lCol: PCol;
  sameWidth: Boolean;
begin
  // If the count of columns is equal to the max colcount of the file format
  // then it is likely that dummy columns have been added -> delete all empty
  // columns (starting at the right) until the first non-empty column is found
  if AWorksheet.Cols.Count =  SizeInt(FLimitations.MaxColCount) then
  begin
    c := AWorksheet.Cols.Count - 1;
    lCol := PCol(AWorksheet.Cols[c]);
    w := lCol.Width;
    while c >= 0 do begin
      lCol := PCol(AWorksheet.Cols[c]);
      if not SameValue(lCol^.Width, w, EPS) then
        break;
      if AWorksheet.FindNextCellInCol(0, c) <> nil then
        break;
      AWorksheet.RemoveCol(c);
      dec(c);
    end;
  end;

  if AWorksheet.Cols.Count < 2 then
    exit;

  // Check whether all columns have the same column width
  sameWidth := true;
  w := PCol(AWorksheet.Cols[0])^.Width;
  for c := 1 to AWorksheet.Cols.Count-1 do begin
    lCol := PCol(AWorksheet.Cols[c]);
    if not SameValue(lCol^.Width, w, EPS) then
    begin
      sameWidth := false;
      break;
    end;
  end;

  if sameWidth then begin
    // At this point we know that all columns have the same width. We pass this
    // to the DefaultColWidth ...
    AWorksheet.WriteDefaultColWidth(w, FWorkbook.Units);

    // ...and delete all column records with non-default format
    for c := AWorksheet.Cols.Count-1 downto 0 do begin
      lCol := PCol(AWorksheet.Cols[c]);
      if lCol^.FormatIndex = 0 then AWorksheet.RemoveCol(c);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This procedure checks whether all rows have the same height and removes the
  row records if they do. Such unnecessary row records are often written
  when an Office application converts a file to another format.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadReader.FixRows(AWorkSheet: TsWorksheet);
const
  EPS = 1E-3;
var
  r: Cardinal;
  h: Single;
  lRow: PRow;
begin
  if AWorksheet.Rows.Count <= 1 then
    exit;

  // Check whether all rows have the same height
  h := PRow(AWorksheet.Rows[0])^.Height;
  for r := 1 to AWorksheet.Rows.Count-1 do begin
    lRow := PRow(AWorksheet.Rows[r]);
    if not SameValue(lRow^.Height, h, EPS) then
      exit;
  end;

  // At this point we know that all rows have the same height. We pass this
  // to the DefaultRowHeight ...
  AWorksheet.WriteDefaultRowHeight(h, FWorkbook.Units);

  // ... and delete all row records with non-default format.
  for r := AWorksheet.Rows.Count-1 downto 0 do begin
    lRow := PRow(AWorksheet.Rows[r]);
    if lRow^.FormatIndex = 0 then AWorksheet.RemoveRow(r);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Default file reading method.

  Opens the file and calls ReadFromStream. Data are stored in the workbook
  specified during construction.

  @param  AFileName The input file name.
  @see    TsWorkbook
-------------------------------------------------------------------------------}
procedure TsCustomSpreadReader.ReadFromFile(AFileName: string;
  AParams: TsStreamParams = []);
var
  stream, fs: TStream;
begin
  if (boFileStream in Workbook.Options) then
    stream := TFileStream.Create(AFileName, fmOpenRead + fmShareDenyNone)
  else
  if (boBufStream in Workbook.Options) then
    stream := TBufStream.Create(AFileName, fmOpenRead + fmShareDenyNone)
  else
  begin
    stream := TMemoryStream.Create;
    fs := TFileStream.Create(AFilename, fmOpenRead + fmShareDenyNone);
    try
      (stream as TMemoryStream).CopyFrom(fs, fs.Size);
      stream.Position := 0;
    finally
      fs.Free;
    end;
  end;

  try
    ReadFromStream(stream, AParams);
  finally
    stream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This routine has the purpose to read the workbook data from the stream.
  It should be overriden in descendent classes.

  Its basic implementation here assumes that the stream is a TStringStream and
  the data are provided by calling ReadFromStrings. This mechanism is valid
  for wikitables.

  Data will be stored in the workbook defined at construction.

  @param  AData     Workbook which is filled by the data from the stream.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadReader.ReadFromStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  AStringStream: TStringStream;
  AStrings: TStringList;
begin
  AStringStream := TStringStream.Create('');
  AStrings := TStringList.Create;
  try
    AStringStream.CopyFrom(AStream, AStream.Size);
    AStringStream.Seek(0, soFromBeginning);
    AStrings.Text := AStringStream.DataString;
    ReadFromStrings(AStrings, AParams);
  finally
    AStringStream.Free;
    AStrings.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads workbook data from a string list. This abstract implementation does
  nothing and raises an exception. Must be overridden, like for wikitables.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadReader.ReadFromStrings(AStrings: TStrings;
  AParams: TsStreamParams = []);
begin
  Unused(AStrings, AParams);
  raise Exception.Create(rsUnsupportedReadFormat);
end;


{------------------------------------------------------------------------------}
{                             TsCustomSpreadWriter                             }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the writer. Has the workbook to be written as a parameter to
  apply the localization information found in its FormatSettings.
  Creates an internal instance of the number format list according to the
  file format being read/written.

  @param AWorkbook  Workbook from with the file is written. This parameter is
                    passed from the workbook which creates the writer.
-------------------------------------------------------------------------------}
constructor TsCustomSpreadWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  // Number formats
  FNumFormatList := TStringList.Create;
  AddBuiltinNumFormats;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the writer.
  Destroys the internal number format list.
-------------------------------------------------------------------------------}
destructor TsCustomSpreadWriter.Destroy;
begin
  FreeAndNil(FNumFormatList);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in number formats to the NumFormatList

  The method has to be overridden because the descendants know the special
  requirements of the file format.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.AddBuiltinNumFormats;
begin
  // to be overridden by descendents
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format string is already contained in the
  the writer's internal number format list. If yes, the list index is returned.
-------------------------------------------------------------------------------}
function TsCustomSpreadWriter.FindNumFormatInList(ANumFormatStr: String): Integer;
begin
  for Result:=0 to FNumFormatList.Count-1 do
    if SameText(ANumFormatStr, FNumFormatList[Result]) then
      exit;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  If formatting features of a cell are not supported by the destination file
  format of the writer, here is the place to apply replacements.
  Must be overridden by descendants, nothin happens here. See BIFF2.

  @param  ACell  Pointer to the cell being investigated. Note that this cell
                 does not belong to the workbook, but is a cell of the
                 FFormattingStyles array.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.FixFormat(ACell: PCell);
begin
  Unused(ACell);
  // to be overridden
end;

{@@ ----------------------------------------------------------------------------
  Determines the size of the worksheet to be written. VirtualMode is respected.
  Is called when the writer needs the size for output. Column and row count
  limitations are repsected as well.

  @param   AWorksheet  Worksheet to be written
  @param   AFirsRow    Index of first row to be written
  @param   ALastRow    Index of last row
  @param   AFirstCol   Index of first column to be written
  @param   ALastCol    Index of last column to be written
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.GetSheetDimensions(AWorksheet: TsWorksheet;
  out AFirstRow, ALastRow, AFirstCol, ALastCol: Cardinal);
begin
  if (boVirtualMode in AWorksheet.Workbook.Options) then
  begin
    AFirstRow := 0;
    AFirstCol := 0;
    ALastRow := LongInt(AWorksheet.VirtualRowCount)-1;
    ALastCol := LongInt(AWorksheet.VirtualColCount)-1;
  end else
  begin
    Workbook.UpdateCaches;
    AFirstRow := AWorksheet.GetFirstRowIndex;
    if AFirstRow = Cardinal(-1) then
      AFirstRow := 0;  // this happens if the sheet is empty and does not contain row records
    AFirstCol := AWorksheet.GetFirstColIndex;
    if AFirstCol = Cardinal(-1) then
      AFirstCol := 0;  // this happens if the sheet is empty and does not contain col records
    ALastRow := AWorksheet.GetLastRowIndex;
    ALastCol := AWorksheet.GetLastColIndex;
  end;
  if AFirstCol >= Limitations.MaxColCount then
    AFirstCol := Limitations.MaxColCount-1;
  if AFirstRow >= Limitations.MaxRowCount then
    AFirstRow := Limitations.MaxRowCount-1;
  if ALastCol >= Limitations.MaxColCount then
    ALastCol := Limitations.MaxColCount-1;
  if ALastRow >= Limitations.MaxRowCount then
    ALastRow := Limitations.MaxRowCount-1;
end;

{@@ ----------------------------------------------------------------------------
  Copies the format strings from the workbook's NumFormatList to the writer's
  internal NumFormatList.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.ListAllNumFormats;
var
  i: Integer;
  numFmt: TsNumFormatParams;
  numFmtStr: String;
begin
  for i:=0 to Workbook.GetNumberFormatCount - 1 do
  begin
    numFmt := Workbook.GetNumberFormat(i);
    if numFmt <> nil then
    begin
      numFmtStr := numFmt.NumFormatStr;
      if FindNumFormatInList(numFmtStr) = -1 then
        FNumFormatList.Add(numFmtStr);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Helper function for the spreadsheet writers. Writes the cell value to the
  stream. Calls the WriteNumber method of the worksheet for writing a number,
  the WriteDateTime method for writing a date/time etc.

  @param  ACell    Pointer to the worksheet cell being written
  @param  AStream  Stream to which data are written

  @see    TsCustomSpreadWriter.WriteCellsToStream
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.WriteCellToStream(AStream: TStream; ACell: PCell);
begin
  if HasFormula(ACell) then
    WriteFormula(AStream, ACell^.Row, ACell^.Col, ACell)
  else
    case ACell.ContentType of
      cctBool:
        WriteBool(AStream, ACell^.Row, ACell^.Col, ACell^.BoolValue, ACell);
      cctDateTime:
        WriteDateTime(AStream, ACell^.Row, ACell^.Col, ACell^.DateTimeValue, ACell);
      cctEmpty:
        WriteBlank(AStream, ACell^.Row, ACell^.Col, ACell);
      cctError:
        WriteError(AStream, ACell^.Row, ACell^.Col, ACell^.ErrorValue, ACell);
      cctNumber:
        WriteNumber(AStream, ACell^.Row, ACell^.Col, ACell^.NumberValue, ACell);
      cctUTF8String:
        WriteLabel(AStream, ACell^.Row, ACell^.Col, ACell^.UTF8StringValue, ACell);
    end;

  if FWorksheet.ReadComment(ACell) <> '' then
    WriteComment(AStream, ACell);
end;

{@@ ----------------------------------------------------------------------------
  Helper function for the spreadsheet writers.

  Iterates all cells on a list, calling the appropriate write method for them.

  @param  AStream The output stream.
  @param  ACells  List of cells to be writeen
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.WriteCellsToStream(AStream: TStream;
  ACells: TsCells);
var
  cell: PCell;
begin
  for cell in ACells do
    WriteCellToStream(AStream, cell);
end;

{@@ ----------------------------------------------------------------------------
  (Pseudo-) abstract method writing a cell comment to the stream.
  The cell comment is written immediately after the cell content.
  NOTE: This is not good for XLSX and BIFF8.

  Must be overridden by descendents.

  @param  ACell      Pointer to the cell containing the comment to be written
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.WriteComment(AStream: TStream; ACell: PCell);
begin
  Unused(AStream, ACell);
end;

{@@ ----------------------------------------------------------------------------
  Basic method which is called when writing a formula to a stream. The formula
  is already stored in the cell fields.
  Present implementation does nothing. Needs to be overridden by descendants.

  @param   AStream   Stream to be written
  @param   ARow      Row index of the cell containing the formula
  @param   ACol      Column index of the cell containing the formula
  @param   ACell     Pointer to the cell containing the formula and being written
                     to the stream
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.WriteFormula(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
end;

{@@ ----------------------------------------------------------------------------
  Default file writing method.

  Opens the file and calls WriteToStream
  The workbook written is the one specified in the constructor of the writer.

  @param  AFileName           The output file name.
  @param  AOverwriteExisting  If the file already exists it will be replaced.
  @param  AParams             Optional parameters to control stream access

  @see    TsWorkbook
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean = False; AParams: TsStreamParams = []);
var
  OutputFile: TStream;
  lMode: Word;
begin
  if AOverwriteExisting then
    lMode := fmCreate or fmOpenWrite
  else
    lMode := fmCreate;

  if (boFileStream in FWorkbook.Options) then
    OutputFile := TFileStream.Create(AFileName, lMode)
  else
  if (boBufStream in Workbook.Options) then
    OutputFile := TBufStream.Create(AFileName, lMode)
  else
    OutputFile := TMemoryStream.Create;

  try
    WriteToStream(OutputFile, AParams);
    if OutputFile is TMemoryStream then
      (OutputFile as TMemoryStream).SaveToFile(AFileName);
  finally
    OutputFile.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  This routine has the purpose to write the workbook to a stream.
  Present implementation writes to a stringlists by means of WriteToStrings;
  this behavior is required for wikitables.
  Must be overriden in descendent classes for all other cases.

  @param  AStream   Stream to which the workbook is written
  @param  AParams   Optional parameters to control stream access
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.WriteToStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  list: TStringList;
begin
  list := TStringList.Create;
  try
    WriteToStrings(list, AParams);
    list.SaveToStream(AStream);
  finally
    list.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the worksheet to a list of strings. Not implemented here, needs to
  be overridden by descendants. See wikitables.
-------------------------------------------------------------------------------}
procedure TsCustomSpreadWriter.WriteToStrings(AStrings: TStrings;
  AParams: TsStreamParams = []);
begin
  Unused(AStrings, AParams);
  raise Exception.Create(rsUnsupportedWriteFormat);
end;

type
  TsSpreadFormatData = class
  private
    FFormatID: TsSpreadFormatID;        // Format identifier
    FName: String;                      // Text to be used in FileDialog filter
    FTechnicalName: String;             // Text to be used e.g. in Titlebar
    FFileExtensions: array of String;   // File extensions used by this format
    FReaderClass: TsSpreadReaderClass;  // Class for reading these files
    FWriterClass: TsSpreadWriterClass;  // Class for writing these files
    function GetFileExtension(AIndex: Integer): String;
    function GetFileExtensionCount: Integer;
  public
    constructor Create(AFormatID: TsSpreadFormatID; AReaderClass: TsSpreadReaderClass;
      AWriterClass: TsSpreadWriterClass; AFormatName, ATechnicalName: String;
      const AExtensions: Array of String);
//      ACanReadFromClipboard, ACanWriteToClipboard: Boolean);
    function GetFileFilterMask(ASeparator: Char): String;

//    property CanReadFromClipboard: boolean read FCanReadClipboard;
//    property CanWriteToClipboard: boolean read FCanWriteClipboard;
    property FormatID: TsSpreadFormatID read FFormatID;
    property FormatName: String read FName;
    property FileExtension[AIndex: Integer]: String read GetFileExtension;
    property FileExtensionCount: Integer read GetFileExtensionCount;
    property ReaderClass: TsSpreadReaderClass read FReaderClass;
    property TechnicalName: String read FTechnicalName;
    property WriterClass: TsSpreadWriterClass read FWriterClass;
  end;

  { TsSpreadFormatRegistry }

  TsSpreadFormatRegistry = class
  private
    FList: TFPList;
    FCachedData: TsSpreadFormatData;
    FCachedFormatID: TsSpreadFormatID;
    function GetDefaultExt(AFormatID: TsSpreadFormatID): String;
    function GetFormatName(AFormatID: TsSpreadFormatID): String;
    function GetReaderClass(AFormatID: TsSpreadFormatID): TsSpreadReaderClass;
    function GetTechnicalName(AFormatID: TsSpreadFormatID): String;
    function GetWriterClass(AFormatID: TsSpreadFormatID): TsSpreadWriterClass;
  protected
    function Add(AData: TsSpreadFormatData): Integer;
    function FindFormatID(AFormatID: TsSpreadFormatID): TsSpreadFormatData;
    function IndexOf(AFormatID: TsSpreadFormatID): Integer;
  public
    constructor Create;
    destructor Destroy; override;
    function GetAllSpreadFilesMask(AExtSeparator: Char;
      AFileAccess: TsSpreadFileAccess): String;
    function GetAllExcelFilesMask(AExtSeparator: Char): String;
    function GetFileFilter(AListSeparator, AExtSeparator: Char;
      AFileAccess: TsSpreadFileAccess; const APriorityFormats: array of TsSpreadFormatID;
      AllSpreadFormats: Boolean = false; AllExcelFormats: Boolean = false): String;
    function GetFormatArray(AFileAccess: TsSpreadFileAccess;
      const APriorityFormats: array of TsSpreadFormatID): TsSpreadFormatIDArray;
    function GetFormatArrayFromFileName(AFileAccess: TsSpreadFileAccess;
      const AFileName: String; APriorityFormat: TsSpreadFormatID = sfidUnknown): TsSpreadFormatIDArray;

    property DefaultExt[AFormatID: TsSpreadFormatID]: String read GetDefaultExt;
    property FormatName[AFormatID: TsSpreadFormatID]: String read GetFormatName;
    property ReaderClass[AFormatID: TsSpreadFormatID]: TsSpreadReaderClass read GetReaderClass;
    property TechnicalName[AFormatID: TsSpreadFormatID]: String read GetTechnicalName;
    property WriterClass[AFormatID: TsSpreadFormatID]: TsSpreadWriterClass read GetWriterClass;
  end;

var
  SpreadFormatRegistry: TsSpreadFormatRegistry;

{==============================================================================}
{                           TsSpreadFormatData                                 }
{==============================================================================}

constructor TsSpreadFormatData.Create(AFormatID: TsSpreadFormatID;
  AReaderClass: TsSpreadReaderClass; AWriterClass: TsSpreadWriterClass;
  AFormatName, ATechnicalName: String; const AExtensions: array of String);
var
  i: Integer;
begin
  FFormatID := AFormatID;
  FReaderClass := AReaderClass;
  FWriterClass := AWriterClass;
  FName := AFormatName;
  FTechnicalName := ATechnicalName;
  SetLength(FFileExtensions, Length(AExtensions));
  for i:=0 to High(FFileExtensions) do FFileExtensions[i] := AExtensions[i];
end;

function TsSpreadFormatData.GetFileExtension(AIndex: Integer): String;
begin
  Result := FFileExtensions[AIndex];
end;

function TsSpreadFormatData.GetFileExtensionCount: Integer;
begin
  Result := Length(FFileExtensions);
end;

function TsSpreadFormatData.GetFileFilterMask(ASeparator: Char): String;
var
  i: Integer;
begin
  Result := '*' + FFileExtensions[0];
  for i:= 1 to High(FFileExtensions) do
    Result := Result + ASeparator + '*' + FFileExtensions[i];
end;


{==============================================================================}
{                         TsSpreadFormatRegistry                               }
{==============================================================================}

constructor TsSpreadFormatRegistry.Create;
begin
  inherited;
  FList := TFPList.Create;
  FCachedFormatID := sfidUnknown;
  FCachedData := nil;
end;

destructor TsSpreadFormatRegistry.Destroy;
var
  i: Integer;
begin
  for i := FList.Count-1 downto 0 do TObject(FList[i]).Free;
  FList.Free;

  inherited;
end;

function TsSpreadFormatRegistry.Add(AData: TsSpreadFormatData): Integer;
begin
  Result := FList.Add(AData);
end;

function TsSpreadFormatRegistry.FindFormatID(AFormatID: TsSpreadFormatID): TsSpreadFormatData;
var
  idx: Integer;
begin
  if AFormatID <> FCachedFormatID then
  begin
    idx := IndexOf(AFormatID);
    if idx = -1 then
    begin
      FCachedData := nil;
      FCachedFormatID := sfidUnknown;
    end else
    begin
      FCachedData := TsSpreadFormatData(FList[idx]);
      FCachedFormatID := AFormatID;
    end;
  end;
  Result := FCachedData;
end;

function TsSpreadFormatRegistry.GetDefaultExt(AFormatID: TsSpreadFormatID): String;
var
  data: TsSpreadFormatData;
begin
  data := FindFormatID(AFormatID);
  if data <> nil then
    Result := data.FileExtension[0] else
    Result := '';
end;

function TsSpreadFormatRegistry.GetAllSpreadFilesMask(AExtSeparator: Char;
  AFileAccess: TsSpreadFileAccess): String;
var
  L: TStrings;
  data: TsSpreadFormatData;
  ext: String;
  i, j: Integer;
begin
  Result := '';
  L := TStringList.Create;
  try
    for i:=0 to FList.Count-1 do
    begin
      data := TsSpreadFormatData(FList[i]);
      case AFileAccess of
        faRead  : if data.ReaderClass = nil then continue;
        faWrite : if data.WriterClass = nil then continue;
      end;
      for j:=0 to data.FileExtensionCount-1 do
      begin
        ext := data.FileExtension[j];
        if L.IndexOf(ext) = -1 then
          L.Add(ext);
      end;
    end;
    if L.Count > 0 then
    begin
      Result := '*' + L[0];
      for i := 1 to L.Count-1 do
        Result := Result + AExtSeparator + '*' + L[i];
    end;
  finally
    L.Free;
  end;
end;

function TsSpreadFormatRegistry.GetAllExcelFilesMask(AExtSeparator: Char): String;
var
  j: Integer;
  L: TStrings;
  data: TsSpreadFormatData;
  ext: String;
begin
  L := TStringList.Create;
  try
    // good old BIFF...
    if (IndexOf(ord(sfExcel8)) <> -1) or
       (IndexOf(ord(sfExcel5)) <> -1) or
       (IndexOf(ord(sfExcel2)) <> -1) then L.Add('*.xls');

    // Excel 2007+
    j := IndexOf(ord(sfOOXML));
    if j <> -1 then
    begin
      data := TsSpreadFormatData(FList[j]);
      for j:=0 to data.FileExtensionCount-1 do
      begin
        ext := data.FileExtension[j];
        if L.IndexOf(ext) = -1 then
          L.Add('*' + ext);
      end;
    end;

    L.Delimiter := AExtSeparator;
    L.StrictDelimiter := true;
    Result := L.DelimitedText;
  finally
    L.Free;
  end;
end;

function TsSpreadFormatRegistry.GetFileFilter(AListSeparator, AExtSeparator: Char;
  AFileAccess: TsSpreadFileAccess; const APriorityFormats: array of TsSpreadFormatID;
  AllSpreadFormats: Boolean = false; AllExcelFormats: Boolean = false): String;
var
  i, idx: Integer;
  L: TStrings;
  s: String;
  data: TsSpreadFormatData;
begin
  // Bring the formats listed in APriorityFormats to the top
  if Length(APriorityFormats) > 0 then
    for i := High(APriorityFormats) downto Low(APriorityFormats) do
    begin
      idx := IndexOf(APriorityFormats[i]);
      data := TsSpreadFormatData(FList[idx]);
      FList.Delete(idx);
      FList.Insert(0, data);
    end;

  L := TStringList.Create;
  try
    L.Delimiter := AListSeparator;
    L.StrictDelimiter := true;
    if AllSpreadFormats then
    begin
      s := GetAllSpreadFilesMask(AExtSeparator, AFileAccess);
      if s <> '' then
      begin
        L.Add(rsAllSpreadsheetFiles);
        L.Add(GetAllSpreadFilesMask(AExtSeparator, AFileAccess));
      end;
    end;
    if AllExcelFormats then
    begin
      s := GetAllExcelFilesMask(AExtSeparator);
      if s <> '' then
      begin
        L.Add(Format('%s (%s)', [rsAllExcelFiles, s]));
        L.Add(s);
      end;
    end;
    for i:=0 to FList.Count-1 do
    begin
      data := TsSpreadFormatData(FList[i]);
      case AFileAccess of
        faRead  : if data.ReaderClass = nil then Continue;
        faWrite : if data.WriterClass = nil then Continue;
      end;
      s := data.GetFileFilterMask(AExtSeparator);
      L.Add(Format('%s %s (%s)', [data.FormatName, rsFiles, s]));
      L.Add(s);
    end;
    Result := L.DelimitedText;
  finally
    L.Free;
  end;
end;

function TsSpreadFormatRegistry.GetFormatArray(AFileAccess: TsSpreadFileAccess;
  const APriorityFormats: array of TsSpreadFormatID): TsSpreadFormatIDArray;
var
  i, n, idx: Integer;
  data: TsSpreadFormatData;
begin
  // Rearrange the formats such the one noted in APriorityFormats are at the top
  if Length(APriorityFormats) > 0 then
    for i := High(APriorityFormats) downto Low(APriorityFormats) do
    begin
      idx := IndexOf(APriorityFormats[i]);
      data := TsSpreadFormatData(FList[idx]);
      FList.Delete(idx);
      FList.Insert(0, data);
    end;

  SetLength(Result, FList.Count);
  n := 0;
  for i := 0 to FList.Count-1 do
  begin
    data := TsSpreadFormatData(FList[i]);
    case AFileAccess of
      faRead  : if data.ReaderClass = nil then Continue;
      faWrite : if data.WriterClass = nil then Continue;
    end;
    Result[n] := data.FormatID;
    inc(n);
  end;
  SetLength(Result, n);
end;

function TsSpreadFormatRegistry.GetFormatArrayFromFileName(
  AFileAccess: TsSpreadFileAccess; const AFileName: String;
  APriorityFormat: TsSpreadFormatID = sfidUnknown): TsSpreadFormatIDArray;
var
  idx: Integer;
  i, j, n: Integer;
  ext: String;
  data: TsSpreadFormatData;
begin
  ext := Lowercase(ExtractFileExt(AFileName));

  if APriorityFormat <> sfidUnknown then
  begin
    // Bring the priority format to the top
    idx := IndexOf(APriorityFormat);
    FList.Exchange(0, idx);
  end;

  SetLength(Result, FList.Count);
  n := 0;
  for i := 0 to FList.Count - 1 do
  begin
    data := TsSpreadFormatData(FList[i]);
    case AFileAccess of
      faRead  : if data.ReaderClass = nil then Continue;
      faWrite : if data.WriterClass = nil then Continue;
    end;
    for j:=0 to data.FileExtensionCount-1 do
      if Lowercase(data.FileExtension[j]) = ext then
      begin
        Result[n] := data.FormatID;
        inc(n);
      end;
  end;


  SetLength(Result, n);

  if APriorityFormat <> sfidUnknown then
    // Restore original order
    FList.Exchange(idx, 0);
end;

function TsSpreadFormatRegistry.GetFormatName(AFormatID: TsSpreadFormatID): String;
var
  data: TsSpreadFormatData;
begin
  data := FindFormatID(AFormatID);
  if data <> nil then
    Result := data.FormatName else
    Result := '';
end;

function TsSpreadFormatRegistry.GetReaderClass(AFormatID: TsSpreadFormatID): TsSpreadReaderClass;
var
  data: TsSpreadFormatData;
begin
  data := FindFormatID(AFormatID);
  if data <> nil then
    Result := data.ReaderClass else
    Result := nil;
end;

function TsSpreadFormatRegistry.GetTechnicalName(AFormatID: TsSpreadFormatID): String;
var
  data: TsSpreadFormatData;
begin
  data := FindFormatID(AFormatID);
  if data <> nil then
    Result := data.TechnicalName else
    Result := '';
end;

function TsSpreadFormatRegistry.GetWriterClass(AFormatID: TsSpreadFormatID): TsSpreadWriterClass;
var
  data: TsSpreadFormatData;
begin
  data := FindFormatID(AFormatID);
  if data <> nil then
    Result := data.WriterClass else
    Result := nil;
end;

function TsSpreadFormatRegistry.IndexOf(AFormatID: TsSpreadFormatID): Integer;
begin
  for Result := 0 to FList.Count - 1 do
    if TsSpreadFormatData(FList[Result]).FormatID = AFormatID then
      exit;
  Result := -1;
end;


{==============================================================================}
{                         Public utility functions                             }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Registers a new reader/writer pair for a given spreadsheet file format

  AFormat identifies the file format, see sfXXXX declarations in built-in
  fpstypes.

  The system is open to user-defined formats. In this case, AFormat must have
  the value "sfUser". The format identifier is calculated as a negative number,
  stored in the TsSpreadFormatData class and returned as function result.
  This value is needed when calling fpspreadsheet's ReadFromXXXX and WriteToXXXX
  methods to specify the file format.
-------------------------------------------------------------------------------}
function RegisterSpreadFormat(AFormat: TsSpreadsheetFormat;
  AReaderClass: TsSpreadReaderClass; AWriterClass: TsSpreadWriterClass;
  AFormatName, ATechnicalName: String; const AFileExtensions: array of String): TsSpreadFormatID;
var
  fmt: TsSpreadFormatData;
  n: Integer;
begin
  if AFormat <> sfUser then begin
    n := SpreadFormatRegistry.IndexOf(ord(AFormat));
    if n >= 0 then
      raise Exception.Create('[RegisterSpreadFormat] Spreadsheet format is already registered.');
  end;

  if Length(AFileExtensions) = 0 then
    raise Exception.Create('[RegisterSpreadFormat] File extensions needed for registering a file format.');

  if (AFormatName = '') or (ATechnicalName = '') then
    raise Exception.Create('[RegisterSpreadFormat] File format name is not specified.');

  fmt := TsSpreadFormatData.Create(ord(AFormat), AReaderClass, AWriterClass,
    AFormatName, ATechnicalName, AFileExtensions);
  n := SpreadFormatRegistry.Add(fmt);
  if (AFormat = sfUser) then
  begin
    if (n <= ord(sfUser)) then n := n + ord(sfUser) + 1;
    fmt.FFormatID := -n;
  end;
  Result := fmt.FormatID;
end;

function GetFileFormatFilter(AListSeparator, AExtSeparator: Char;
  AFileAccess: TsSpreadFileAccess; const APriorityFormats: array of TsSpreadFormatID;
  AllSpreadFormats: Boolean = false; AllExcelFormats: Boolean = false): String;
begin
  Result := SpreadFormatRegistry.GetFileFilter(AListSeparator, AExtSeparator,
    AFileAccess, APriorityFormats, AllSpreadFormats, AllExcelFormats);
end;

function GetSpreadFormats(AFileAccess: TsSpreadFileAccess;
  const APriorityFormats: array of TsSpreadFormatID): TsSpreadFormatIDArray;
begin
  Result := SpreadFormatRegistry.GetFormatArray(AFileAccess, APriorityFormats);
end;

function GetSpreadFormatsFromFileName(
  AFileAccess: TsSpreadFileAccess; AFileName: TFileName;
  APriorityFormat: TsSpreadFormatID = sfidUnknown): TsSpreadFormatIDArray;
begin
  Result := SpreadFormatRegistry.GetFormatArrayFromFileName(
    AFileAccess, AFileName, APriorityFormat);
end;

function GetSpreadFormatExt(AFormatID: TsSpreadFormatID): String;
begin
  Result := SpreadFormatRegistry.DefaultExt[AFormatID];
end;

function GetSpreadFormatName(AFormatID: TsSpreadFormatID): String;
begin
  Result := SpreadFormatRegistry.FormatName[AFormatID];
end;

function GetSpreadTechnicalName(AFormatID: TsSpreadFormatID): String;
begin
  Result := SpreadFormatRegistry.TechnicalName[AFormatID];
end;

function GetSpreadReaderClass(AFormatID: TsSpreadFormatID): TsSpreadReaderClass;
begin
  Result := SpreadFormatRegistry.ReaderClass[AFormatID];
end;

function GetSpreadWriterClass(AFormatID: TsSpreadFormatID): TsSpreadWriterClass;
begin
  Result := SpreadFormatRegistry.WriterClass[AFormatID];
end;


initialization
  SpreadFormatRegistry := TsSpreadFormatRegistry.Create;

finalization
  SpreadFormatRegistry.Free;


end.
