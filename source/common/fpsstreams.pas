unit fpsStreams;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

uses
  SysUtils, Classes;

var
  DEFAULT_STREAM_BUFFER_SIZE: Integer = 1024 * 1024;       // 1 MB

type
  { A buffered stream }
  TBufStream = class(TStream)
  private
    FFileStream: TFileStream;
    FMemoryStream: TMemoryStream;
    FFileStreamPos: Int64;
    FFileStreamSize: Int64;
    FBufWritten: Boolean;
    FBufSize: Int64;
    FKeepTmpFile: Boolean;
    FFileName: String;
    FFileMode: Word;
  protected
    procedure CreateFileStream;
    function GetPosition: Int64; override;
    function GetSize: Int64; override;
    class function IsWritingMode(AMode: Word): Boolean;
    procedure SetSize64(const NewValue: Int64); override;
  public
    constructor Create(AFileName: String; AMode: Word;
      ABufSize: Cardinal = Cardinal(-1)); overload;
    constructor Create(ATempFile: String; AKeepFile: Boolean = false;
      ABufSize: Cardinal = Cardinal(-1)); overload;
    constructor Create(ABufSize: Cardinal = Cardinal(-1)); overload;
    destructor Destroy; override;
    procedure Clear;
    procedure FillBuffer;
    procedure FlushBuffer;
    function Read(var Buffer; Count: Longint): Longint; override;
    function Seek(const Offset: Int64; Origin: TSeekOrigin): Int64; override;
    function Write(const ABuffer; ACount: Longint): Longint; override;
  end;

procedure ResetStream(var AStream: TStream);


implementation

uses
  Math;

{ Resets the stream position to the beginning of the stream. }
procedure ResetStream(var AStream: TStream);
begin
  if AStream <> nil then
    AStream.Position := 0;
end;


{==============================================================================}
{                               TBufStream                                     }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Constructor of the TBufStream. Creates a memory stream and prepares everything
  to create also a file stream if the stream size exceeds ABufSize bytes.

  @param  ATempFile   File name for the file stream. If an empty string is
                      used a temporary file name is created by calling GetTempFileName.
  @param  AKeepFile   If true and the stream is in WritingMode the stream is
                      flushed to file when the stream is
                      destroyed. If false the file is deleted when the stream
                      is destroyed.
  @param  ABufSize    Maximum size of the memory stream before swapping to file
                      starts. Value is given in bytes.
-------------------------------------------------------------------------------}
constructor TBufStream.Create(ATempFile: String; AKeepFile: Boolean = false;
  ABufSize: Cardinal = Cardinal(-1));
begin
  if ATempFile = '' then
    ATempFile := ChangeFileExt(GetTempFileName, '.~abc');
  // Change extension because of naming conflict if the name of the main file
  // is determined by GetTempFileName also. Happens in internaltests suite.
  FFileName := ATempFile;
  FKeepTmpFile := AKeepFile;
  FMemoryStream := TMemoryStream.Create;
  // The file stream is only created when needed because of possible conflicts
  // of random file names.
  if ABufSize = Cardinal(-1) then
    FBufSize := DEFAULT_STREAM_BUFFER_SIZE
  else
    FBufSize := ABufSize;
  FFileMode := fmCreate + fmOpenRead;
end;

{@@
  Constructor of the TBufStream. Creates a memory stream and prepares everything
  to create also a file stream if the streamsize exceeds ABufSize bytes. The
  stream created by this constructor is mainly intended to serve a temporary
  purpose, it is not stored permanently to file.

  @param  ABufSize    Maximum size of the memory stream before swapping to file
                      starts. Value is given in bytes.
}
constructor TBufStream.Create(ABufSize: Cardinal = Cardinal(-1));
begin
  Create('', false, ABufSize);
end;

{@@
  Constructor of the TBufStream. When swapping to file it will create a file
  stream using the given file mode. This kind of BufStream is considered as a
  fast replacement of TFileStream.

  @param  AFileName   File name for the file stream. If an empty string is
                      used a temporary file name is created by calling GetTempFileName.
  @param  AMode       FileMode for the file stream (fmCreate, fmOpenRead etc.)
  @param  ABufSize    Maximum size of the memory stream before swapping to file
                      starts. Value is given in bytes.
}
constructor TBufStream.Create(AFileName: String; AMode: Word;
  ABufSize: Cardinal = Cardinal(-1));
var
  keep: Boolean;
begin
  keep := IsWritingMode(AMode);
  Create(AFileName, keep, ABufSize);
  FFileMode := AMode;
end;

destructor TBufStream.Destroy;
begin
  // Write current buffer content to file
  if FKeepTmpFile then FlushBuffer;

  // Free streams and delete temporary file, if requested
  FreeAndNil(FMemoryStream);
  FreeAndNil(FFileStream);
  if not FKeepTmpFile and (FFileName <> '') and IsWritingMode(FFileMode) then
    DeleteFile(FFileName);

  inherited Destroy;
end;

{ Creation of the file stream is delayed because of naming conflicts of other
  streams are needed with random file names as well (the files do not yet exist
  when the streams are created and therefore get the same name by GetTempFileName! }
procedure TBufStream.CreateFileStream;
begin
  if FFileStream = nil then begin
    if FFileName = '' then FFileName := ChangeFileExt(GetTempFileName, '.~abc');
    FFileStream := TFileStream.Create(FFileName, FFileMode);
    FFileStreamSize := FFileStream.Size;
    FFileStreamPos := 0;
  end;
end;

{ Reads FBufSize bytes from the stream into the buffer.
  Called when reading. }
procedure TBufStream.FillBuffer;
var
  p, n: Int64;
begin
  p := GetPosition;
  FMemoryStream.Clear;
  FMemoryStream.Position := 0;
  FFileStream.Position := p;
  n := Min(FBufSize, FFileStreamSize - p);
  FMemoryStream.CopyFrom(FFileStream, n);
  FMemoryStream.Position := 0;
  FFileStream.Position := p;  // The file stream ends where the memorystream begins!
  FFileStreamPos := p;
end;

{ Flushes the contents of the memory stream to file
  Called when writing. }
procedure TBufStream.FlushBuffer;
begin
  if (FMemoryStream.Size > 0) and not FBufWritten and IsWritingMode(FFileMode) then
  begin
    FMemoryStream.Position := 0;
    CreateFileStream;
    FFileStream.CopyFrom(FMemoryStream, FMemoryStream.Size);
    FFileStreamPos := FFileStream.Position;
    FFileStreamSize := FFileStream.Size;
    FMemoryStream.Clear;
    FBufWritten := true;
  end;
end;

{ Returns the buffer position. This is the buffer position of the bytes written
  to file, plus the current position in the memory buffer }
function TBufStream.GetPosition: Int64;
begin
  if FFileStream = nil then
    Result := FMemoryStream.Position
  else
  //  Result := FFileStream.Position + FMemoryStream.Position;
    Result := FFileStreamPos + FMemoryStream.Position;
end;

{ Returns the size of the stream. Both memory and file streams are considered
  if needed. }
function TBufStream.GetSize: Int64;
var
  n: Int64;
begin
  if IsWritingMode(FFileMode) then begin
    if FFileStream <> nil then
      n := FFileStreamSize
//      n := FFileStream.Size
    else
      n := 0;
    if n = 0 then n := FMemoryStream.Size;
    Result := Max(n, GetPosition);
  end else begin
    CreateFileStream;
    Result := FFileStreamSize;
  end;
end;

{@@
  Returns true if the stream is in WritingMode.
  "WritingMode" means that the stream is primarily used for writing. The
  memory stream is initially empty but fills during writing, it is written to
  disk when it is full.
  The (unnamend) opposite of "WritingMode" indicates that the stream is used
  for reading. The memory stream is initially full, but the stream pointer is at
  it start. When data are read the stream pointer advances towards the end.
  When the requested data are not contained in the memory stream another
  ABufSize of bytes are read into the memory stream. }
class function TBufStream.IsWritingMode(AMode: Word): Boolean;
begin
  Result := (AMode and (fmCreate or fmOpenWrite or fmOpenReadWrite) <> 0);
end;

{@@
  Reads a given number of bytes into a buffer and return the number of bytes
  read. If the bytes are not in the memory stream they are read from the file
  stream.

  @param  Buffer  Buffer into which the bytes are read. Sufficient space must
                  have been allocated for Count bytes
  @param  Count   Number of bytes to read from the stream
  @return Number of bytes that were read from the stream.}
function TBufStream.Read(var Buffer; Count: Longint): Longint;
var
  p: Int64;
begin
  p := GetPosition;  // Save stream position

  // Case 1: Memory stream is empty
  if FMemoryStream.Size = 0 then begin
    CreateFileStream;
    if IsWritingMode(FFileMode) then begin
      Result := FFileStream.Read(Buffer, Count);
      FFileStreamPos := FFileStream.Position;
    end else begin
      FillBuffer;
      Result := FMemoryStream.Read(Buffer, Count);
    end;
    exit;
  end;

  // Case 2: All "Count" bytes are contained in memory stream starting at current position
  if FMemoryStream.Position + Count <= FMemoryStream.Size then begin
    Result := FMemoryStream.Read(Buffer, Count);
    exit;
  end;

  // Case 3: Memory stream is not empty but contains only part of the bytes requested
  if IsWritingMode(FFileMode) then begin
    FlushBuffer;
    FFileStream.Position := p;
    Result := FFileStream.Read(Buffer, Count);
    FFileStreamPos := p + Count;
  end else begin
    FillBuffer;
    Result := FMemoryStream.Read(Buffer, Count);
  end;
end;

function TBufStream.Seek(const Offset: Int64; Origin: TSeekOrigin): Int64;
var
  oldPos: Int64;
  newPos: Int64;
begin
  oldPos := GetPosition;
  case Origin of
    soBeginning : newPos := Offset;
    soCurrent   : newPos := oldPos + Offset;
    soEnd       : newPos := GetSize - Offset;
  end;

  // case #1: New position is within buffer, no file stream yet
  if (FFileStream = nil) and (newPos < FMemoryStream.Size) then
  begin
    FMemoryStream.Position := newPos;
    Result := FMemoryStream.Position;
    exit;
  end;

  CreateFileStream;

  // case #2: New position is within buffer, file stream exists
//  if (newPos >= FFileStream.Position) and (newPos < FFileStream.Position + FMemoryStream.Size)
  if (newPos >= FFileStreamPos) and (newPos < FFileStreamPos + FMemoryStream.Size)
  then begin
  //  FMemoryStream.Position := newPos - FFileStream.Position;
    FMemoryStream.Position := newPos - FFileStreamPos;
    Result := newpos; //FMemoryStream.Position;
    exit;
  end;

  // case #3: New position is outside buffer
  if IsWritingMode(FFileMode) then
    FlushBuffer;
  FFileStream.Position := newPos;
  FFileStreamPos := newPos;
  FMemoryStream.Position := 0;
  if not IsWritingMode(FFileMode) then
    FillBuffer;
end;

procedure TBufStream.SetSize64(const NewValue: Int64);
begin
  if NewValue = 0 then
    Clear
  else
    raise Exception.Create('Setting the TBufStream.Size is not allowed.');
end;

procedure TBufStream.Clear;
begin
  FMemoryStream.Clear;
  if not Assigned(FFileStream) then
    CreateFileStream;
  FFileStream.Size := 0;
  FFileStream.Position := 0;
  FFileStreamPos := 0;
end;

function TBufStream.Write(const ABuffer; ACount: LongInt): LongInt;
var
  savedPos: Int64;
begin
  // Case #1: Bytes fit into buffer
  if FMemoryStream.Position + ACount < FBufSize then
  begin
    Result := FMemoryStream.Write(ABuffer, ACount);
    FBufWritten := false;
  end else
  // Case #2: Buffer would overflow
  begin;
    savedPos := GetPosition;
    if (FMemorystream.Size = 0) and (ACount > 0) and (FFileStream = nil) then
      CreateFileStream;
    FlushBuffer;
    FFileStream.Position := savedPos;
    Result := FFileStream.Write(ABuffer, ACount);
    FFileStreamPos := savedPos + ACount;
    FFileStreamSize := FFileStream.Size;
  end;
end;


end.
