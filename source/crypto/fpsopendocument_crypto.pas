unit fpsOpenDocument_Crypto;

{$MODE ObjFPC}{$H+}
{.$DEFINE UNZIP_ABBREVIA}     // Remove this define when zipper is fixed.

interface

uses
  Classes, SysUtils,
  {$IFDEF UNZIP_ABBREVIA}
  ABUnzper,
  {$ENDIF}
  fpsTypes, fpsOpenDocument;

type
  TsSpreadOpenDocReaderCrypto = class(TsSpreadOpenDocReader)
  private
    function CalcPasswordHash(ADecryptionInfo: TsOpenDocManifestFileEntry;
      APassword: String): TBytes;
  protected
    function Decrypt(AStream: TStream; ADecryptionInfo: TsOpenDocManifestFileEntry;
      APassword: String; ADestStream: TStream; out AErrorMsg: String): Boolean; override;
    function SupportsDecryption: Boolean; override;
    {$IFDEF UNZIP_ABBREVIA}
    function UnzipToStream(AStream: TStream; AZippedFile: String; ADestStream: TStream): Boolean; override;
    {$ENDIF}
  end;

var
  sfidOpenDocument_Crypto: TsSpreadFormatID;

implementation

uses
  zStream,
  fpsReaderWriter, fpsCryptoProc;

{ Decompresses the source stream and stored the output in the destination stream. }
procedure Decompress(ASrcStream, ADestStream: TStream);
var
  decompressor: TDecompressionStream;
begin
  decompressor := TDecompressionStream.Create(ASrcStream, true);
  try
    ADestStream.CopyFrom(decompressor, 0);  // 0 --> entire src stream
  finally
    decompressor.Free;
  end;
end;

{-------------------------------------------------------------------------------
                         TsSpreadOpenDocReaderCrypto
-------------------------------------------------------------------------------}

{ AStream contains one encrypted xml file of the ods file structure. The method
  decrypts the stream based on the information provided in ADecryptionInfo and
  using the given (unhashed) user password. The output is stored in the
  destination stream. }
function TsSpreadOpenDocReaderCrypto.Decrypt(AStream: TStream;
  ADecryptionInfo: TsOpenDocManifestFileEntry; APassword: String;
  ADestStream: TStream; out AErrorMsg: String): Boolean;
var
  pwdHash: TBytes;
  iv: TBytes;
  tmpStream: TStream;
  algorithm: String;
begin
  Result := false;

  algorithm := LowerCase(ADecryptionInfo.AlgorithmName);
  if (algorithm = 'aes128-cbc') or (algorithm='aes192-cbc') or (algorithm='aes256-cbc') then
    algorithm := 'aes'
  else
    algorithm := '';

  if algorithm = '' then
    exit;

  tmpStream := TMemoryStream.Create;
  try
    // Calculated password hash
    pwdHash := CalcPasswordHash(ADecryptionInfo, APassword);

    // Decrypt
    iv := DecodeBase64(ADecryptionInfo.InitializationVector);
    case algorithm of
      'aes':
        AErrorMsg := Decrypt_AES_CBC(pwdHash[0], Length(pwdHash)*8, @iv[0], AStream, tmpStream);
      else
        AErrorMsg := 'Encryption method not supported.';
    end;
    if (AErrorMsg <> '') then
      exit;

    // Verify decrypted (but still compressed) stream
    // OpenDocument-v1.2-part3, section 3.8.3: "The digest is build from the compressed unencrypted file"
    tmpStream.Position := 0;
    if not VerifyDecrypt(tmpStream, ADecryptionInfo.EncryptionData_CheckSum, ADecryptionInfo.EncryptionData_ChecksumType) then
      AErrorMsg := 'Checksum error';
    if AErrorMsg <> '' then
      exit;

    // Decompress the decrypted stream
    Decompress(tmpStream, ADestStream);
    ADestStream.Position := 0;

    // Success!
    Result := true;
  finally
    tmpStream.Free;
  end;
end;

{ Calculates the hash value of the user-provided passwort.
  Hash creation is determined by information stored in ADecryptionInfo. }
function TsSpreadOpenDocReaderCrypto.CalcPasswordHash(
  ADecryptionInfo: TsOpenDocManifestFileEntry;
  APassword: String): TBytes;
var
  pwdHash: TBytes;
  salt: TBytes;
  numIterations: Integer;
  keySize: Integer;
begin
  Result := nil;

  // Generate start key
  case LowerCase(ADecryptionInfo.StartKeyGenerationName) of
    'sha256': pwdHash := Calc_SHA256(APassword[1], Length(APassword));
  else
    raise EFpSpreadsheetReader.Create('Unsupported start key generator ' + ADecryptionInfo.StartKeyGenerationName);
  end;

  // Generate derived key
  numIterations := ADecryptionInfo.IterationCount;
  keySize := ADecryptionInfo.KeySize;
  salt := DecodeBase64(ADecryptionInfo.Salt);
  if LowerCase(ADecryptionInfo.KeyDerivationName) = 'pbkdf2' then
    Result := PBKDF2_HMAC_SHA1(pwdHash, salt, numIterations, keySize)
  else
    raise EFpSpreadsheetReader.Create('Unsupported key generation method ' + ADecryptionInfo.KeyDerivationName);
end;


{ Tells the calling routine that this reader is able to decrypt ods files. }
function TsSpreadOpenDocReaderCrypto.SupportsDecryption: Boolean;
begin
  Result := true;
end;


{$IFDEF UNZIP_ABBREVIA}
{ Extracts the specified file from the compressed stream (AStream) to the
  ADestStream.
  Uses the ABBREVIA library for this purpose (because FCL Stripper fails to
  extract the encrypted file). }
function TsSpreadOpenDocReaderCrypto.UnzipToStream(AStream: TStream;
  AZippedFile: String; ADestStream: TStream): Boolean;
var
  unzipper: TABUnzipper;
begin
  Result := false;
  unzipper := TABUnzipper.Create(nil);
  try
    unzipper.Stream := AStream;
    try
      unzipper.ExtractToStream(AZippedFile, ADestStream);
      ADestStream.Position := 0;
      Result := true;
    except
      raise;
    end;
  finally
    unzipper.Free;
  end;
end;
{$ENDIF}


{==============================================================================}
                             initialization
{==============================================================================}

{ Registers this reader for fpSpreadsheet }

  sfidOpenDocument_Crypto := RegisterSpreadFormat(sfUser,
    TsSpreadOpenDocReaderCrypto, nil,
    STR_FILEFORMAT_OPENDOCUMENT, 'ODS', [STR_OPENDOCUMENT_CALC_EXTENSION]
  );

end.

