unit xlsxdecrypter;
{
  Some of the ideas are aquired from http://www.lyquidity.com/devblog/?p=35
  (the `internal` or `default password`): VelvetSweatshop
}
{$ifdef fpc}
  {$mode objfpc}{$H+}
//  {$mode delphi}
{$endif}

interface

uses
  Classes
  , SysUtils
  , sha1
  , DCPrijndael
  ;

  const
    CFB_Signature = $E11AB1A1E011CFD0; // Compound File Binary Signature
                                       // Weird is the documentation is equal to
                                       // $D0CF11E0A1B11AE1, but here is inversed
                                       // maybe related to litle endian thing?!!

    // EncryptionHeaderFlags as defined in 2.3.1 [MS-OFFCRYPTO]
    ehfAES       = $00000004;
    //ehfExternal  = $00000008;
    //ehfDocProps  = $00000010;
    ehfCryptoAPI = $00000020;

    // AlgorithmID
    algRC4    = $00006801;
    algAES128 = $0000660E;
    algAES192 = $0000660F;
    algAES256 = $00006610;

    // HashID
    hsSHA1    = $00008004;

    // ProviderType
    prRC4     = $00000001;
    prAES     = $00000018;


  type

    TVersion = packed record
      Major : Word;
      Minor : Word
    end;

    { Defined in Section 2.3.2, 2.3.4.5   [MS-OFFCRYPTO] }
    TEncryptionHeader = record
      Flags       : DWord;   { defined in section 2.3.1 [MS-OFFCRYPTO] }
      SizeExtra   : DWord;   { Must be equal to 0 }
      AlgorithmID : DWord;   { $00006801 -- RC4   }
                             { $0000660E -- AES128}
                             { $0000660F -- AES192}
                             { $00006610 -- AES256}

      HashID      : DWord;   { $00008004 -- SHA1  }

      KeySize     : DWord;   { RC4    -- 40bits to 128bits (8-bit increments) }
                             { AES128 -- 128 bits }
                             { AES192 -- 192 bits }
                             { AES256 -- 256 bits }

      ProviderType: DWord;   { $00000001 -- RC4   }
                             { $00000018 -- AES   }

      Reserved1   : DWord;   { Ignored }
      Reserved2   : DWord;   { Must be equal to 0 }
      CSP_Name    : string;
    end;

    { Defined in Section 2.3.3  [MS-OFFCRYPTO] }
    TEncryptionVerifier = record
      SaltSize             : DWord;
      Salt                 : array[0..15] of Byte;
      EncryptedVerifier    : array[0..15] of Byte;
      VerifierHashSize     : DWord;
      EncryptedVerifierHash: array[0..31] of Byte; // RC4 needs only 20 bytes
    end;

    // The EncryptionInfo Stream as define in 2.3.4.5 [MS-OFFCRYPTO]
    TEncryptionInfo = record
      Version   : TVersion;
      Flags     : DWord;
      HeaderSize: DWord;
      Header    : TEncryptionHeader;
      Verifier  : TEncryptionVerifier;
    end;

  { TExcelFileDecryptor }
  TExcelFileDecryptor = class
  private
    FEncInfo : TEncryptionInfo;
    FEncryptionKey : TBytes;

    // return empty string if everything done right otherwise the error message.
    function InitEncryptionInfo(AStream: TStream): string;

    //CheckPasswordInternal should be called after InitEncryptionInfo
    function CheckPasswordInternal( APassword: UnicodeString ): Boolean;

  public
    // return empty string if everything done right otherwise the error message.
    function Decrypt(inFileName: string; outStream: TStream): string; overload;
    function Decrypt(inStream: TStream; outStream: TStream):string; overload;

    // made this private because I don't know if it'll work with other passwords
    function Decrypt(inFileName: string; outStream: TStream; APassword: UnicodeString): string; overload;
    function Decrypt(inStream: TStream; outStream: TStream; APassword: UnicodeString): string; overload;

    // return true if the password is correct.
    function CheckPassword(AFileName: string; APassword: UnicodeString): Boolean;
    function CheckPassword(AStream: TStream; APassword: UnicodeString): Boolean;

    function isEncryptedAndSupported(AFileName: string): Boolean;
    function isEncryptedAndSupported(AStream: TStream): Boolean;
  end;



implementation

uses
  fpolebasic
  ;

procedure ConcatToByteArray(var outArray: TBytes; Arr1: TBytes; Arr2: TBytes);
var
  LenArr1 : Integer;
  LenArr2 : Integer;
begin
  LenArr1 := Length(Arr1);
  LenArr2 := Length(Arr2);

  SetLength( outArray, LenArr1 + LenArr2 );

  if LenArr1 > 0 then
    Move(Arr1[0], outArray[0], LenArr1);

  if LenArr2 > 0 then
    Move(Arr2[0], outArray[LenArr1], LenArr2);
end;

procedure ConcatToByteArray(var outArray: TBytes; AValue: DWord; Arr: TBytes);
var
  LenArr : Integer;
begin
  LenArr := Length(Arr);

  SetLength( outArray, 4 + LenArr );

  Move(AValue, outArray[0], 4);

  if LenArr > 0 then
    Move(Arr[0], outArray[4], LenArr);
end;

procedure ConcatToByteArray(var outArray: TBytes; Arr: TBytes; AValue: DWord);
var
  LenArr : Integer;
begin
  LenArr := Length(Arr);

  SetLength( outArray, 4 + LenArr );

  if LenArr > 0 then
    Move(Arr[0], outArray[0], LenArr);

  Move(AValue, outArray[LenArr], 4);
end;

function TExcelFileDecryptor.InitEncryptionInfo(AStream: TStream): string;
var
  EncInfoStream: TMemoryStream;
  OLEStorage: TOLEStorage;
  OLEDocument: TOLEDocument;
  FileSignature: QWord;
  Pos : Int64;

  Err : string;
begin
  Err := '';

  if not Assigned(AStream) then
    Exit( 'Stream is null' );

  AStream.Position := 0;
  FileSignature := AStream.ReadQWord;
  if FileSignature <> QWord(CFB_Signature) then
    Exit( 'Wrong file signature' );

  EncInfoStream := TMemoryStream.Create;
  try
    OLEStorage := TOLEStorage.Create;
    try
      OLEDocument.Stream := EncInfoStream;
      AStream.Position := 0;
      OLEStorage.ReadOLEStream(AStream, OLEDocument, 'EncryptionInfo');
      if OLEDocument.Stream.Size = 0 then
        raise Exception.Create('EncryptionInfo stream not found.');

      EncInfoStream.Position := 0;

      { Major Version: $0002 = Excel 2003
                       $0003 = Excel 2007 | 2007 SP1
                       $0004 = Excel 2007 SP2 (not sure about 2010 | 2013) }
      FEncInfo.Version.Major := EncInfoStream.ReadWord;
      if (FEncInfo.Version.Major <> 3) and (FEncInfo.Version.Major <> 4) then
        raise Exception.Create('File must be created with 2007 or 2010');

      { Minor Version: must be equal to $0002 }
      FEncInfo.Version.Minor := EncInfoStream.ReadWord;
      if FEncInfo.Version.Minor <> 2  then
         raise Exception.Create('Incorrect File Version');

      FEncInfo.Flags         := EncInfoStream.ReadDWord;
      FEncInfo.HeaderSize    := EncInfoStream.ReadDWord;

      ///
      /// ENCRYPTION HEADER
      ///
      Pos := EncInfoStream.Position;
      With FEncInfo.Header do
      begin
        Flags       := EncInfoStream.ReadDWord;
        if (Flags and ehfCryptoAPI) <> ehfCryptoAPI then
          raise Exception.Create('File not encrypted');
        if (Flags and ehfAES) <> ehfAES then
          raise Exception.Create('Encryption must be AES');

        SizeExtra   := EncInfoStream.ReadDWord;
        if SizeExtra <> 0 then
          raise Exception.Create('Wrong Header.SizeExtra');

        AlgorithmID := EncInfoStream.ReadDWord;
        if   ( AlgorithmID <> algAES128 )
          and( AlgorithmID <> algAES192 )
          and( AlgorithmID <> algAES256 )
          //and( AlgorithmID <> algRC4    ) // not used by ECMA-376 format
          then
            raise Exception.Create('Unknown Encryption Algorithm');

        HashID      := EncInfoStream.ReadDWord;
        if HashID <> hsSHA1 then
          raise Exception.Create('Unknown Hashing Algorithm');

        KeySize     := EncInfoStream.ReadDWord;
        if ( (AlgorithmID = algAES128) and (KeySize <> 128) )
         or( (AlgorithmID = algAES192) and (KeySize <> 192) )
         or( (AlgorithmID = algAES256) and (KeySize <> 256) )
         //or( (AlgorithmID = algRC4) and (KeySize < 40 or KeySize > 128) )
         then
           raise Exception.Create('Incorrect Key Size');

        ProviderType:= EncInfoStream.ReadDWord;
        if ( ProviderType <> prAES )
          //and( FEncInfo.Header.ProviderType <> prRC4 )
          then
            raise Exception.Create('Unknown Provider');

        Reserved1   := EncInfoStream.ReadDWord;
        Reserved2   := EncInfoStream.ReadDWord;
        if Reserved2 <> 0 then
          raise Exception.Create('Reserved2 must equal to 0');

        //CSP_Name    := Not needed
        // CSP: Should be Microsoft Enhanced RSA and AES Cryptographic Provider
        //  or  Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)
        //Skip CSP Name
        EncInfoStream.Position := Pos + FEncInfo.HeaderSize;
      end;

      ///
      /// ENCRYPTION VERIFIER
      ///
      with FEncInfo.Verifier do
      begin
        SaltSize       := EncInfoStream.ReadDWord;
        if FEncInfo.Verifier.SaltSize <> 16 then
          raise Exception.Create('Incorrect salt size');

        EncInfoStream.ReadBuffer(Salt[0], SaltSize);
        EncInfoStream.ReadBuffer(EncryptedVerifier[0], SaltSize);

        VerifierHashSize   := EncInfoStream.ReadDWord;

        if FEncInfo.Header.ProviderType = prAES then
          EncInfoStream.ReadBuffer( EncryptedVerifierHash[0], 32);
        { for RC4
        else if FEncInfo.Header.ProviderType = prRC4 then
          EncInfoStream.ReadBuffer( EncryptedVerifierHash[0], 20); }
      end;

      Err := '';
    except
      on E: Exception do
        Err := E.Message;
    end;
  finally
    if Assigned(OLEStorage) then
      OLEStorage.Free;

    EncInfoStream.Free;
  end;

  Result := Err;
end;

function TExcelFileDecryptor.CheckPasswordInternal(APassword: UnicodeString): Boolean;
var
  AES_Cipher: TDCP_rijndael;

  ConcArr : TBytes;
  LastHash: TSHA1Digest;

  Iterator, i: DWord;

  X1_Buff: array[0..63] of byte;
  X2_Buff: array[0..63] of byte;
  X1_Hash: TSHA1Digest;
  X2_Hash: TSHA1Digest;

  EncryptionKeySize : Integer;

  Verifier  : array[0..15] of Byte;
  VerifierHash: array[0..31] of Byte;// Needs only 20bytes to hold the SHA1
                                     // but needs 32bytes to hold the decrypted hash
begin
  // if no password used, use microsoft default.
  if APassword = '' then
    APassword := 'VelvetSweatshop';

  //// [MS-OFFCRYPTO]
  //// 2.3.4.7 ECMA-376 Document Encryption Key Generation

  // 1.1.Concat Salt and Password
  //     Calculate SHA1(0) =  SHA1(salt + password)
  ConcatToByteArray( ConcArr
                   , FEncInfo.Verifier.Salt
                   , TEncoding.Unicode.GetBytes(APassword));
  LastHash := SHA1Buffer( ConcArr[0], Length(ConcArr) );

  // 1.2.Calculate SHA1(n) = SHA1(iterator + SHA1(n-1) ) -- iterator is 32bit
  for  Iterator := 0 to 49999 do
  begin
    ConcatToByteArray(ConcArr, Iterator, LastHash);
    LastHash := SHA1Buffer( ConcArr[0], Length(ConcArr) );
  end;

  // 1.3.Claculate final hash, SHA1(final) = SHA1(H(n) + block) -- block = 0 (32bit)
  ConcatToByteArray(ConcArr, LastHash, 0);
  LastHash := SHA1Buffer( ConcArr[0], Length(ConcArr) );


  //// 2.Derive the encryption key.
  // 2.1 cbRequiredKeyLength for AES is 128,192,256bit ?!!! must be < 40bytes
  // 2.2 cbHash = 20bytes ( 160bit),, length of SHA1 hash
  // 2.3 + 2.4 Claculate X1 and X2 the SHA of the generated 64bit Arrays.

  // FillByte(X1_Buff[0], 64, $36);
  // FillByte(X2_Buff[0], 64, $5C);
  for i := 0 to 19 do
  begin
    X1_Buff[i] := LastHash[i] xor $36;
    X2_Buff[i] := LastHash[i] xor $5C;
  end;
  for i := 20 to 63 do
  begin
    X1_Buff[i] := $36;
    X2_Buff[i] := $5C;
  end;

  X1_Hash := SHA1Buffer( X1_Buff[0], Length(X1_Buff) );
  X2_Hash := SHA1Buffer( X2_Buff[0], Length(X2_Buff) );

  // 2.5 Concat X1, X2 -> X3 = X1 + X2 (X3 = 40bytes in length)
  //ConcatToByteArray( ConcArr, X1_Hash, X2_Hash );

  // 2.6 Let keyDerived be equal to the first cbRequiredKeyLength bytes of X3.
  //     We'll fill the Encryption key on the fly, so we won't need X3
  //     This Key (FEncryptionKey) is used for decryption method
  EncryptionKeySize := FEncInfo.Header.KeySize div 8; // Convert Size from bits to bytes
  SetLength(FEncryptionKey, EncryptionKeySize);
  if EncryptionKeySize <= 20 then
  begin
    Move(X1_Hash[0], FEncryptionKey[0], EncryptionKeySize);
  end
  else
  begin
    Move(X1_Hash[0], FEncryptionKey[0], EncryptionKeySize);
    Move(X2_Hash[0], FEncryptionKey[20], EncryptionKeySize-20);
  end;

  //// 2.3.4.9 Password Verification
  // 1. Encryption key is FEncryptionKey

  // 2. Decrypt the EncryptedVerifier
  AES_Cipher := TDCP_rijndael.Create(nil);
  AES_Cipher.Init( FEncryptionKey[0], FEncInfo.Header.KeySize, nil );
  AES_Cipher.DecryptECB(FEncInfo.Verifier.EncryptedVerifier[0] , Verifier[0]);

  // 3. Decrypt the DecryptedVerifierHash
  AES_Cipher.Burn;
  AES_Cipher.Init( FEncryptionKey[0], FEncInfo.Header.KeySize, nil );
  AES_Cipher.DecryptECB(FEncInfo.Verifier.EncryptedVerifierHash[0] , VerifierHash[0]);
  AES_Cipher.DecryptECB(FEncInfo.Verifier.EncryptedVerifierHash[16], VerifierHash[16]);
  AES_Cipher.Free;

  // 4. Calculate SHA1(Verifier)
  LastHash := SHA1Buffer(Verifier[0], Length(Verifier));

  // 5. Compare results
  Result := (CompareByte( LastHash[0], VerifierHash[0], 20) = 0);
end;

function TExcelFileDecryptor.Decrypt(inFileName: string; outStream: TStream
  ): string;
begin
  Result := Decrypt(inFileName, outStream, 'VelvetSweatshop' );
end;

function TExcelFileDecryptor.Decrypt(inFileName: string; outStream: TStream;
  APassword: UnicodeString): string;
Var
  inStream : TFileStream;
begin
  if not FileExists(inFileName) then
    Exit( inFileName + ' not found.' );

  try
    inStream := TFileStream.Create( inFileName, fmOpenRead );

    inStream.Position := 0;
    Result := Decrypt( inStream, outStream, APassword );
  finally
    inStream.Free;
  end;
end;

function TExcelFileDecryptor.Decrypt(inStream: TStream; outStream: TStream
  ): string;
begin
  Result := Decrypt(inStream, outStream, 'VelvetSweatshop' );
end;

function TExcelFileDecryptor.Decrypt(inStream: TStream; outStream: TStream;
  APassword: UnicodeString): string;
var
  OLEStream: TMemoryStream;
  OLEStorage: TOLEStorage;
  OLEDocument: TOLEDocument;

  AES_Cipher :  TDCP_rijndael;
  inData  : TBytes;
  outData : TBytes;
  StreamSize : QWord;
  KeySizeByte: Integer;

  Err : string;
begin
  if (not Assigned(inStream)) or (not Assigned(outStream)) then
    Exit( 'streams must be assigned' );

  Err := InitEncryptionInfo(inStream);
  if Err  <> '' then
    Exit( 'Error when initializing Encryption Info'#10#13 + Err );

  if not CheckPasswordInternal(APassword) then
    Exit( 'Wrong password' );

  // read the encoded stream into memory
  OLEStream := TMemoryStream.Create;
  try
    OLEStorage := TOLEStorage.Create;
    try
      OLEDocument.Stream := OLEStream;
      inStream.Position := 0;
      OLEStorage.ReadOLEStream(inStream, OLEDocument, 'EncryptedPackage');
      if OLEDocument.Stream.Size = 0 then
        raise Exception.Create('EncryptedPackage stream not found.');

      // Start decryption
      OLEStream.Position:=0;
      outStream.Position:=0;

      StreamSize := OLEStream.ReadQWord;

      KeySizeByte := FEncInfo.Header.KeySize div 8;
      SetLength(inData, KeySizeByte);
      SetLength(outData, KeySizeByte);

      AES_Cipher := TDCP_rijndael.Create(nil);
      AES_Cipher.Init( FEncryptionKey[0], FEncInfo.Header.KeySize, nil );

      While StreamSize > 0 do
      begin
        OLEStream.ReadBuffer(inData[0], KeySizeByte);
        AES_Cipher.DecryptECB(inData[0], outData[0]);

        if StreamSize < KeySizeByte then
          outStream.WriteBuffer(outData[0], StreamSize) // Last block less then key size
        else
          outStream.WriteBuffer(outData[0], KeySizeByte);

        if StreamSize < KeySizeByte then
           StreamSize := 0
        else
          Dec(StreamSize, KeySizeByte);
      end;

      AES_Cipher.Free;

       /////
    except
      Err := 'EncryptedPackage not found';
    end;
  finally
    if Assigned(OLEStorage) then
      OLEStorage.Free;

    OLEStream.Free;
  end;
  Exit( Err );
end;

function TExcelFileDecryptor.isEncryptedAndSupported(AFileName: string
  ): Boolean;
var
  AStream : TStream;
begin
  if not FileExists(AFileName) then
    Exit( False );

  try
    AStream := TFileStream.Create( AFileName, fmOpenRead );

    AStream.Position := 0;
    //FStream.CopyFrom(AStream, AStream.Size);

    Result := isEncryptedAndSupported( AStream );
  finally
    AStream.Free;
  end;
end;

function TExcelFileDecryptor.isEncryptedAndSupported(AStream: TStream
  ): Boolean;
begin
  if not Assigned(AStream) then
    Exit( False );

  if InitEncryptionInfo(AStream) <> '' then
    Exit( False );

  Result := True;
end;

function TExcelFileDecryptor.CheckPassword(AFileName: string;
  APassword: UnicodeString): Boolean;
var
  AStream : TStream;
begin
  if not FileExists(AFileName) then
    Exit( False );

  try
    AStream := TFileStream.Create( AFileName, fmOpenRead );

    AStream.Position := 0;

    Result := CheckPassword( AStream, APassword );
  finally
    AStream.Free;
  end;
end;

function TExcelFileDecryptor.CheckPassword(AStream: TStream;
  APassword: UnicodeString): Boolean;
begin
  if not Assigned(AStream) then
    Exit( False );

  AStream.Position := 0;
  if InitEncryptionInfo(AStream) <> '' then
    Exit( False );

  Result := CheckPasswordInternal(APassword);
end;

end.

