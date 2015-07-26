{ fpsxmlcommon.pas
  Unit shared by all xml-type reader/writer classes }

unit fpsxmlcommon;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  laz2_xmlread, laz2_DOM,
  fpSpreadsheet, fpsreaderwriter;

type
  TsSpreadXMLReader = class(TsCustomSpreadReader)
  protected
    procedure ReadXMLFile(out ADoc: TXMLDocument; AFileName: String);
    procedure ReadXMLStream(out ADoc: TXMLDocument; AStream: TStream);
  end;

function GetAttrValue(ANode : TDOMNode; AAttrName : string) : string;
function GetNodeValue(ANode: TDOMNode): String;

procedure UnzipFile(AZipFileName, AZippedFile, ADestFolder: String);
function UnzipToStream(AZipStream: TStream; const AZippedFile: String;
  ADestStream: TStream): Boolean;

implementation

uses
 {$IF FPC_FULLVERSION >= 20701}
  zipper,
 {$ELSE}
  fpszipper,
 {$ENDIF}
  fpsStreams;

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
  if ANode = nil then
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


{------------------------------------------------------------------------------}
{                                 Unzipping                                    }
{------------------------------------------------------------------------------}
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
  FSuccess := True;
  AStream := FOutputStream;
end;

procedure TStreamUnzipper.DoneStream(Sender: TObject; var AStream: TStream;
  AItem: TFullZipFileEntry);
begin
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

end.

