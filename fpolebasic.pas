{
fpolestorage.pas

Writes an OLE document using the OLE virtual layer.

Note: Compatibility with previous version (fpolestorage.pas).
}
unit fpolebasic;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  uvirtuallayer_ole;

type

  { Describes an OLE Document }

  TOLEDocument = record
    // Information about the document
    Stream: TStream;
//    Stream: TMemoryStream;
  end;


  { TOLEStorage }

  TOLEStorage = class
  private
  public
    procedure WriteOLEFile(AFileName: string; AOLEDocument: TOLEDocument; const AOverwriteExisting: Boolean = False; const AStreamName: String='Book');
    procedure WriteOLEStream(AStream: TStream; AOLEDocument: TOLEDocument; const AStreamName: String='Book');
    procedure ReadOLEFile(AFileName: string; AOLEDocument: TOLEDocument; const AStreamName: String='Book');
    procedure ReadOLEStream(AStream: TStream; AOLEDocument: TOLEDocument; const AStreamName: String='Book');
    procedure FreeOLEDocumentData(AOLEDocument: TOLEDocument);
  end;

implementation

uses
  fpsStrings;

{@@ ----------------------------------------------------------------------------
  Writes the OLE document specified in AOLEDocument
  to the file with name AFileName. The routine will fail
  if the file already exists, or if the directory where
  it should be placed doesn't exist.
-------------------------------------------------------------------------------}
procedure TOLEStorage.WriteOLEFile(AFileName: string;
  AOLEDocument: TOLEDocument; const AOverwriteExisting: Boolean;
  const AStreamName: String = 'Book');
var
  RealFile: TFileStream;
begin
  if FileExists(AFileName) then
  begin
    if AOverwriteExisting then
      DeleteFile(AFileName)
      // In Ubuntu it seems that fmCreate does not erase an existing file.
      // Therefore, we delete it manually
    else
      raise EStreamError.CreateFmt(rsFileAlreadyExists, [AFileName]);
  end;

  RealFile := TFileStream.Create(AFileName, fmCreate);
  try
    WriteOLEStream(RealFile, AOLEDocument, AStreamName);
  finally
    RealFile.Free;
  end;
end;

procedure TOLEStorage.WriteOLEStream(AStream: TStream; AOLEDocument: TOLEDocument;
  const AStreamName: String = 'Book');
var
  fsOLE: TVirtualLayer_OLE;
  VLAbsolutePath: String;
  OLEStream: TStream;
  tmpStream: TStream;  // workaround to compiler bug, see bug 22370
begin
  VLAbsolutePath := '/' + AStreamName;   // Virtual layer always uses absolute paths
  fsOLE := TVirtualLayer_OLE.Create(AStream);
  try
    fsOLE.Format;  // Initialize and format the OLE container;
    OLEStream := fsOLE.CreateStream(VLAbsolutePath, fmCreate);
    try
      // woraround for bug 22370
      tmpStream := AOLEDocument.Stream;
      tmpStream.Position := 0;  // Ensures that stream is at the beginning
      // previous code:  AOLEDocument.Stream.Position := 0;
      OLEStream.CopyFrom(tmpStream, tmpStream.Size);
    finally
      OLEStream.Free;
    end;
  finally
    fsOLE.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads an OLE file.
-------------------------------------------------------------------------------}
procedure TOLEStorage.ReadOLEFile(AFileName: string;
  AOLEDocument: TOLEDocument; const AStreamName: String = 'Book');
var
  RealFile: TFileStream;
begin
  RealFile := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
  try
    ReadOLEStream(RealFile, AOLEDocument, AStreamName);
  finally
    RealFile.Free;
  end;
end;


procedure TOLEStorage.ReadOLEStream(AStream: TStream; AOLEDocument: TOLEDocument;
  const AStreamName: String = 'Book');
var
  fsOLE: TVirtualLayer_OLE;
  OLEStream: TStream;
  VLAbsolutePath: UTF8String;
begin
  VLAbsolutePath := '/' + AStreamName; //Virtual layer always use absolute paths.
  fsOLE := TVirtualLayer_OLE.Create(AStream);
  try
    fsOLE.Initialize(); //Initialize the OLE container.
    OLEStream := fsOLE.CreateStream(VLAbsolutePath, fmOpenRead);
    try
      if Assigned(OLEStream) then begin
        if not Assigned(AOLEDocument.Stream) then
          AOLEDocument.Stream := TMemoryStream.Create else
          (AOLEDocument.Stream as TMemoryStream).Clear;
        AOLEDocument.Stream.CopyFrom(OLEStream, OLEStream.Size);
      end;
    finally
      OLEStream.Free;
    end;
  finally
    fsOLE.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Frees all internal objects storable in a TOLEDocument structure
-------------------------------------------------------------------------------}
procedure TOLEStorage.FreeOLEDocumentData(AOLEDocument: TOLEDocument);
begin
  if Assigned(AOLEDocument.Stream) then FreeAndNil(AOLEDocument.Stream);
end;

end.

