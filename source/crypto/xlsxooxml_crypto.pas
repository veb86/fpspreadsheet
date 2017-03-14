unit xlsxooxml_crypto;

interface

uses
  Classes,
  fpstypes, xlsxooxml, xlsxdecrypter;
  
type
  TsSpreadOOXMLReaderCrypto = class(TsSpreadOOXMLReader)
  public
    procedure ReadFromStream(AStream: TStream; APassword: String = '';
      AParams: TsStreamParams = []); override;
  end;

var
  sfidOOXML_Crypto: TsSpreadFormatID;


implementation

uses
  fpsReaderWriter;

procedure TsSpreadOOXMLReaderCrypto.ReadFromStream(AStream: TStream;
  APassword: String = ''; AParams: TsStreamParams = []);
var
  ExcelDecrypt : TExcelFileDecryptor;
  DecryptedStream: TStream;
begin
  ExcelDecrypt := TExcelFileDecryptor.Create;
  try
    AStream.Position := 0;
    if ExcelDecrypt.isEncryptedAndSupported(AStream) then
    begin
      DecryptedStream := TMemoryStream.Create;
      try
        ExcelDecrypt.Decrypt(AStream, DecryptedStream, APassword);
        // Discard encrypted stream and load decrypted one.
        AStream.Free;
        AStream := TMemoryStream.Create;
        DecryptedStream.Position := 0;
        AStream.CopyFrom(DecryptedStream, DecryptedStream.Size);
        AStream.Position := 0;
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


initialization

  // Registers this reader/writer for fpSpreadsheet
  sfidOOXML_Crypto := RegisterSpreadFormat(sfUser,
    TsSpreadOOXMLReaderCrypto, nil,
    STR_FILEFORMAT_EXCEL_XLSX, 'OOXML', [STR_OOXML_EXCEL_EXTENSION, '.xlsm']
  );

end.

end.
