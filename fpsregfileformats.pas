{@@ ----------------------------------------------------------------------------
  Unit fpsRegFileFormats implements registration of the file formats supported
  by fpspreadsheet.

  AUTHORS: Felipe Monteiro de Carvalho, Reinier Olislagers, Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.

  USAGE:   Each unit implementing a new spreadsheet format must register the
           reader/writer and some specific data by calling "RegisterSpreadFormat".
-------------------------------------------------------------------------------}
unit fpsRegFileFormats;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpstypes, fpsReaderWriter; //fpspreadsheet;

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
  fpsStrings;

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

