{ fpsxmlcommon.pas
  Unit shared by all xml-type reader/writer classes }

unit fpsxmlcommon;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  laz2_xmlread, laz2_DOM,
 {$IF FPC_FULLVERSION >= 20701}
  zipper,
 {$ELSE}
  fpszipper,
 {$ENDIF}
  fpstypes, fpsreaderwriter;

type
  TsSpreadXMLReader = class(TsCustomSpreadReader)
  protected
    procedure ReadXMLFile(out ADoc: TXMLDocument; AFileName: String);
    procedure ReadXMLStream(out ADoc: TXMLDocument; AStream: TStream);
  end;

  TStreamUnzipper = class(TUnzipper)
  private
    FInputStream: TStream;
    FOutputStream: TStream;
    FSuccess: Boolean;
    procedure CloseInputStream(Sender: TObject; var AStream: TStream);
    procedure CreateStream(Sender: TObject; var AStream: TStream;
      AItem: TFullZipFileEntry);
    procedure DoneStream(Sender: TObject; var AStream: TStream;
      AItem: TFullZipFileEntry);
    procedure OpenInputStream(Sender: TObject; var AStream: TStream);
  public
    constructor Create(AInputStream: TStream);
    function UnzipFile(const AZippedFile: string; ADestStream: TStream): Boolean;
  end;


function GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
function GetNodeValue(ANode: TDOMNode): String;

function LineEndingToBR(const AText: String): String;
function UTF8TextToXMLText(AText: string; ProcessLineEndings: Boolean = false): string;
function ValidXMLText(var AText: string; ReplaceSpecialChars: Boolean = true;
  ProcessLineEndings: Boolean = false): Boolean;

procedure UnzipFile(AZipFileName, AZippedFile, ADestFolder: String);
function UnzipToStream(AZipStream: TStream; const AZippedFile: String;
  ADestStream: TStream): Boolean;

function CreateTempStream(AWorkbook: TsBasicWorkbook;
  AFileNameBase: String): TStream;
procedure DestroyTempStream(AStream: TStream);


implementation

uses
  (*
 {$IF FPC_FULLVERSION >= 20701}
  zipper,
 {$ELSE}
  fpszipper,
 {$ENDIF}
 *)
  fpsStreams, fpsUtils;

{------------------------------------------------------------------------------}
{                                 Utilities                                    }
{------------------------------------------------------------------------------}

{ Gets value for the specified attribute of the given node.
  Returns empty string if attribute is not found. }
function GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
var
  i: LongWord;
  Found: Boolean;
begin
  Result := '';
  if (ANode = nil) or (ANode.Attributes = nil) then
    exit;

  Found := false;
  i := 0;
  while not Found and (i < ANode.Attributes.Length) do begin
    if ANode.Attributes.Item[i].NodeName = AAttrName then begin
      Found := true;
      Result := ANode.Attributes.Item[i].NodeValue;
    end;
    inc(i);
  end;
end;

{ Returns the text value of a node. Normally it would be sufficient to call
  "ANode.NodeValue", but since the DOMParser needs to preserve white space
  (for the spaces in date/time formats), we have to go more into detail. }
function GetNodeValue(ANode: TDOMNode): String;
var
  child: TDOMNode;
begin
  Result := '';
  child := ANode.FirstChild;
  if Assigned(child) and (child.NodeName = '#text') then
    Result := child.NodeValue;
end;

{@@ ----------------------------------------------------------------------------
  Replaces LineEnding character(s) by '<br />';
-------------------------------------------------------------------------------}
function LineEndingToBR(const AText: String): String;
var
  i: Integer;
begin
  Result := '';
  i := 1;
  while (i <= Length(AText)) do
  begin
    case AText[i] of
      #13: begin
             Result := Result + '<br />';
             if (i < Length(AText)) and (AText[i+1] = #10) then inc(i);
           end;
      #10: Result := Result + '<br />';
      else Result := Result + AText[i];
    end;
    inc(i);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Converts a string encoded in UTF8 to a string usable in XML. For this purpose,
  some characters must be translated.

  @param   AText               Input string encoded as UTF8
  @param   ProcessLineEndings  If TRUE line ending characters are replaced by
                               their HTML entities (e.g., #10 --> '&#10;'
  @return  String usable in XML with some characters replaced by the HTML codes.
-------------------------------------------------------------------------------}
function UTF8TextToXMLText(AText: string;
  ProcessLineEndings: Boolean = false): string;
var
  Idx: Integer;
  AppoSt: string;
begin
  Result := '';
  idx := 1;
  while idx <= Length(AText) do
  begin
    case AText[Idx] of
      '&': begin
        AppoSt := Copy(AText, Idx, 6);
        if (Pos('&amp;',  AppoSt) = 1) or
           (Pos('&lt;',   AppoSt) = 1) or
           (Pos('&gt;',   AppoSt) = 1) or
           (Pos('&quot;', AppoSt) = 1) or
           (Pos('&apos;', AppoSt) = 1) or
           (Pos('&#37;',  AppoSt) = 1)     // %
        then begin
          //'&' is the first char of a special chat, it must not be converted
          Result := Result + AText[Idx];
        end else begin
          Result := Result + '&amp;';
        end;
      end;
      '<': Result := Result + '&lt;';
      '>': Result := Result + '&gt;';
      '"': Result := Result + '&quot;';
      '''':Result := Result + '&apos;';
      '%': Result := Result + '&#37;';
      #10: if ProcessLineEndings then
             Result := Result + '&#10;' else
             Result := Result + #10;
      #13: if ProcessLineEndings then
             Result := Result + '&#13;' else
             Result := Result + #13;
      {     this breaks multi-line labels in xlsx
      #10: begin
             Result := Result + '<br />';
             if (idx < Length(AText)) and (AText[idx+1] = #13) then inc(idx);
           end;
      #13: begin
             Result := Result + '<br />';
             if (idx < Length(AText)) and (AText[idx+1] = #10) then inc(idx);
           end;
           }
    else
      Result := Result + AText[Idx];
    end;
    inc(idx);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks a string for characters that are not permitted in XML strings.
  The function returns FALSE if a character <#32 is contained (except for
  #9, #10, #13), TRUE otherwise. Invalid characters are replaced by a box symbol.

  If ReplaceSpecialChars is TRUE, some other characters are converted
  to valid HTML codes by calling UTF8TextToXMLText

  @param  AText                String to be checked. Is replaced by valid string.
  @param  ReplaceSpecialChars  Special characters are replaced by their HTML
                               codes (e.g. '>' --> '&gt;')
  @param  ProcessLineEndings   If TRUE line ending characters are replaced by
                               their HTML entities.
  @return FALSE if characters < #32 were replaced, TRUE otherwise.
-------------------------------------------------------------------------------}
function ValidXMLText(var AText: string;
  ReplaceSpecialChars: Boolean = true;
  ProcessLineEndings: Boolean = false): Boolean;
const
  BOX = #$E2#$8E#$95;
var
  i: Integer;
begin
  Result := true;
  for i := Length(AText) downto 1 do
    if (AText[i] < #32) and not (AText[i] in [#9, #10, #13]) then begin
      // Replace invalid character by box symbol
      Delete(AText, i, 1);
      Insert(BOX, AText, i);
      Result := false;
    end;
  if ReplaceSpecialChars then
    AText := UTF8TextToXMLText(AText, ProcessLineEndings);
end;

{------------------------------------------------------------------------------}
{                                 Unzipping                                    }
{------------------------------------------------------------------------------}
(*
type
  TStreamUnzipper = class(TUnzipper)
  private
    FInputStream: TStream;
    FOutputStream: TStream;
    FSuccess: Boolean;
    procedure CloseInputStream(Sender: TObject; var AStream: TStream);
    procedure CreateStream(Sender: TObject; var AStream: TStream;
      AItem: TFullZipFileEntry);
    procedure DoneStream(Sender: TObject; var AStream: TStream;
      AItem: TFullZipFileEntry);
    procedure OpenInputStream(Sender: TObject; var AStream: TStream);
  public
    constructor Create(AInputStream: TStream);
    function UnzipFile(const AZippedFile: string; ADestStream: TStream): Boolean;
  end;
*)
constructor TStreamUnzipper.Create(AInputStream: TStream);
begin
  inherited Create;
  OnCloseInputStream := @CloseInputStream;
  OnCreateStream := @CreateStream;
  OnDoneStream := @DoneStream;
  OnOpenInputStream := @OpenInputStream;
  FInputStream := AInputStream
end;

procedure TStreamUnzipper.CloseInputStream(Sender: TObject; var AStream: TStream);
begin
  AStream := nil;
end;

procedure TStreamUnzipper.CreateStream(Sender: TObject; var AStream: TStream;
  AItem: TFullZipFileEntry);
begin
  Unused(AItem);
  FSuccess := True;
  AStream := FOutputStream;
end;

procedure TStreamUnzipper.DoneStream(Sender: TObject; var AStream: TStream;
  AItem: TFullZipFileEntry);
begin
  Unused(AItem);
  AStream := nil;
end;

procedure TStreamUnzipper.OpenInputStream(Sender: TObject; var AStream: TStream);
begin
  AStream := FInputStream;
end;

function TStreamUnzipper.UnzipFile(const AZippedFile: string;
  ADestStream: TStream): Boolean;
begin
  FOutputStream := ADestStream;
  FSuccess := False;
  Files.Clear;
  Files.Add(AZippedFile);
  UnZipAllFiles;
  Result := FSuccess;
end;

{ We have to use our own ReadXMLFile procedure (there is one in xmlread)
  because we have to preserve spaces in element text for date/time separator.
  As a side-effect we have to skip leading spaces by ourselves. }
procedure TsSpreadXMLReader.ReadXMLFile(out ADoc: TXMLDocument; AFileName: String);
var
  stream: TStream;
begin
  if (boBufStream in Workbook.Options) then
    stream := TBufStream.Create(AFilename, fmOpenRead + fmShareDenyWrite)
  else
    stream := TFileStream.Create(AFileName, fmOpenRead + fmShareDenyWrite);

  try
    ReadXMLStream(ADoc, stream);
  finally
    stream.Free;
  end;
end;

procedure TsSpreadXMLReader.ReadXMLStream(out ADoc: TXMLDocument; AStream: TStream);
var
  parser: TDOMParser;
  src: TXMLInputSource;
begin
  parser := TDOMParser.Create;
  try
    parser.Options.PreserveWhiteSpace := true;    // This preserves spaces!
    src := TXMLInputSource.Create(AStream);
    try
      parser.Parse(src, ADoc);
    finally
      src.Free;
    end;
  finally
    parser.Free;
  end;
end;

procedure UnzipFile(AZipFileName, AZippedFile, ADestFolder: String);
var
  list: TStringList;
  unzip: TUnzipper;
begin
  list := TStringList.Create;
  try
    list.Add(AZippedFile);
    unzip := TUnzipper.Create;
    try
      Unzip.OutputPath := ADestFolder;
      Unzip.UnzipFiles(AZipFileName, list);
    finally
      unzip.Free;
    end;
  finally
    list.Free;
  end;
end;


function UnzipToStream(AZipStream: TStream; const AZippedFile: String;
  ADestStream: TStream): Boolean;
var
  unzip: TStreamUnzipper;
  p: Int64;
begin
  p := ADestStream.Position;
  unzip := TStreamUnzipper.Create(AZipStream);
  try
    Result := unzip.UnzipFile(AZippedFile, ADestStream);
    ADestStream.Position := p;
  finally
    unzip.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a basic stream for storing of the individual files. Depending on
  the set workbook options the stream is created as a memory stream (default),
  buffered stream or file stream.

  In the latter two cases a filename mask is provided to create a temporary
  filename around this mask.
-------------------------------------------------------------------------------}
function CreateTempStream(AWorkbook: TsBasicWorkbook;
  AFilenameBase: String): TStream;
begin
  if boFileStream in AWorkbook.Options then
    Result := TFileStream.Create(GetTempFileName('', AFilenameBase), fmCreate)
  else
  if boBufStream in AWorkbook.Options then
    Result := TBufStream.Create(GetTempFileName('', AFilenameBase))
  else
    Result := TMemoryStream.Create;
end;


procedure DestroyTempStream(AStream: TStream);
var
  fn: String;
begin
  // TMemoryStream and TBufStream need not be considered separately,
  // they destroy everything themselves. Only the TFileStream must delete its
  // temporary file.
  if AStream is TFileStream then
  begin
    fn := TFileStream(AStream).Filename;
    AStream.Free;          // Destroy stream before deleting temp file!
    DeleteFile(fn);        // Otherwise the temp file will not be deleted.
  end else
    AStream.Free;
end;


end.

