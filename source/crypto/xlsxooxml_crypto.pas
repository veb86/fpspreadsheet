unit xlsxooxml_crypto;

{$mode objfpc}{$H+}

interface

uses
  Classes,
  fpstypes, fpsUtils, xlsxooxml, xlsxdecrypter;
  
type
  TsSpreadOOXMLReaderCrypto = class(TsSpreadOOXMLReader)
  private
    FNeedsPassword: Boolean;
  public
    class function CheckFileFormat(AStream: TStream): boolean; override;
    function NeedsPassword(AStream: TStream): Boolean; override;
    procedure ReadFromStream(AStream: TStream; APassword: String = '';
      AParams: TsStreamParams = []); override;
    function SupportsDecryption: Boolean; override;
  end;

var
  sfidOOXML_Crypto: TsSpreadFormatID;


implementation

uses
  fpsReaderWriter;

class function TsSpreadOOXMLReaderCrypto.CheckFileFormat(AStream: TStream): boolean;
begin
  Result := inherited;               // This checks for a normal xlsx format ...
  if not Result then
    Result := IsEncrypted(AStream);  // ... and this for a decrypted one.
end;

function TsSpreadOOXMLReaderCrypto.NeedsPassword(AStream: TStream): Boolean;
begin
  Unused(AStream);
  Result := FNeedsPassword;
end;

procedure TsSpreadOOXMLReaderCrypto.ReadFromStream(AStream: TStream;
  APassword: String = ''; AParams: TsStreamParams = []);
var
  ExcelDecrypt : TExcelFileDecryptor;
  DecryptedStream: TStream;
begin
  FNeedsPassword := false;

  ExcelDecrypt := TExcelFileDecryptor.Create;
  try
    AStream.Position := 0;
    if ExcelDecrypt.isEncryptedAndSupported(AStream) then
    begin
      FNeedsPassword := true;
      CheckPassword(AStream, APassword);
      DecryptedStream := TMemoryStream.Create;
      try
        ExcelDecrypt.Decrypt(AStream, DecryptedStream, UnicodeString(APassword));
        // Discard encrypted stream and load decrypted one.
        AStream.Free;
        AStream := TMemoryStream.Create;
        DecryptedStream.Position := 0;
        AStream.CopyFrom(DecryptedStream, DecryptedStream.Size);
        AStream.Position := 0;
        FNeedsPassword := false;    // AStream is not encrypted any more.
      finally
        DecryptedStream.Free;
      end;
    end;
  finally
    ExcelDecrypt.Free;
    AStream.Position := 0;
  end;

  inherited;
end;

function TsSpreadOOXMLReaderCrypto.SupportsDecryption: Boolean;
begin
  Result := true;
end;


initialization

  // Registers this reader/writer for fpSpreadsheet
  sfidOOXML_Crypto := RegisterSpreadFormat(sfUser,
    TsSpreadOOXMLReaderCrypto, nil,
    STR_FILEFORMAT_EXCEL_XLSX, 'OOXML', [STR_OOXML_EXCEL_EXTENSION, '.xlsm']
  );

end.

end.
