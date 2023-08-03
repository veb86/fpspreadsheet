{-------------------------------------------------------------------------------
                  Generic decryption procedures
--------------------------------------------------------------------------------
  Use one of the defines to select the cryptographic library used.
  - WOLFGANG_EHRHARDT_LIB:
      these units are included in fpspreadsheet directly.
  - DCPCRYPT:
      the package DCPCrypt must be added to the "required packages" of
      fpsSpreadsheet_crypto
-------------------------------------------------------------------------------}

unit fpsCryptoProc;

{$mode objfpc}{$H+}

{ Activate one of the following two defines. }
{$DEFINE WOLFGANG_EHRHARDT_LIB}
{.$DEFINE DCPCRYPT}

{$IF DEFINED(DCPCRYPT) AND DEFINED(WOLFGANG_EHRHARDT_LIB)}
  ERROR: Only a single cryptographic library can be selected.
{$ENDIF}

interface

uses
  Classes, SysUtils,
  Base64, sha1,
  {$IF FPC_FullVersion >= 30300}
  fpHashUtils, fpSHA256,
  {$ELSE}
  fpsHashUtils, fpsSHA256,
  {$ENDIF}
  {$IFDEF WOLFGANG_EHRHARDT_LIB}
  aes_type, aes_cbc, aes_ecb;
  {$ENDIF}
  {$IFDEF DCPCRYPT}
  DCPcrypt2, DCPrijndael;
  {$ENDIF}

function ConcatToByteArray(const InArray1, InArray2: TBytes): TBytes;
procedure ConcatToByteArray(var OutArray: TBytes;
  Ptr1: PByte; ACount1: Integer; Ptr2: PByte; ACount2: Integer);
procedure ConcatToByteArray(var OutArray: TBytes;
  Ptr: PByte; ACount: Integer; const Arr: TBytes);

function DecodeBase64(const AString: String): TBytes;

function Calc_SHA1(const Buf; BufSize: LongWord): TBytes;
function Calc_SHA1(Buf: TBytes): TBytes;

function Calc_SHA256(const Buf; BufSize: LongWord): TBytes;
function Calc_SHA256(Buf: TBytes): TBytes;

function PBKDF2_HMAC_SHA1(pass, salt: TBytes; count, kLen: Integer): TBytes;

function Decrypt_AES_ECB(const Key; KeySizeBits: LongWord;
  const InData; var OutData; DataSize: LongWord): String;

function DecryptStream_AES_ECB(const Key; KeySizeBits: LongWord;
  ASrcStream, ADestStream: TStream; ASrcStreamSize: QWord): String;

function Decrypt_AES_CBC(const Key; KeySizeBits: LongWord; InitVector: Pointer;
  ASrcStream, ADestStream: TStream): String;

function VerifyDecrypt(AStream: TStream; CheckSum, ChecksumType: string): boolean;

implementation

uses
  Math, fpsUtils;

function ConcatToByteArray(const InArray1, InArray2: TBytes): TBytes;
begin
  ConcatToByteArray(Result, @InArray1[0], Length(InArray1), @InArray2[0], Length(InArray2));
end;

procedure ConcatToByteArray(var OutArray: TBytes; Ptr1: PByte; ACount1: Integer;
  Ptr2: PByte; ACount2: Integer);
begin
  SetLength(OutArray, ACount1 + ACount2);
  if ACount1 > 0 then
    Move(Ptr1^, OutArray[0], ACount1);
  if ACount2 > 0 then
    Move(Ptr2^, OutArray[ACount1], ACount2);
end;

procedure ConcatToByteArray(var OutArray: TBytes; Ptr: PByte; ACount: Integer;
  const Arr: TBytes);
begin
  ConcatToByteArray(OutArray, Ptr, ACount, @Arr[0], Length(Arr));
end;

function DecodeBase64(const AString: String): TBytes;
begin
  Result := StringToBytes(DecodeStringBase64(AString));
end;

function Calc_SHA1(const Buf; BufSize: LongWord): TBytes;
var
  digest: TSHA1Digest;
begin
  digest := SHA1Buffer(Buf, BufSize);
  SetLength(Result, Length(digest));
  Move(digest[0], Result[0], Length(digest));
end;

function Calc_SHA1(Buf: TBytes): TBytes;
begin
  Result := Calc_SHA1(Buf[0], Length(Buf));
end;

function Calc_SHA256(const Buf; BufSize: LongWord): TBytes;
var
  sha256: TSHA256;
begin
  sha256.Init;
  sha256.Update(@Buf, BufSize);
  sha256.Final;
  SetLength(Result, SizeOf(TSHA256Digest));
  Move(sha256.Digest[0], Result[0], SizeOf(TSHA256Digest));
end;

function Calc_SHA256(Buf: TBytes): TBytes;
begin
  Result := Calc_SHA256(Buf[0], Length(Buf));
end;

function RPad(Data: TBytes; PadByte: Byte; ALen: Integer): TBytes;
var
  L: Integer;
begin
  L := Length(Data);
  if L < ALen then
  begin
    SetLength(Result, ALen);
    Move(Data[0], Result[0], L);
    FillChar(Result[L], ALen-L, PadByte);
  end else
    Result := Data;
end;

function Fill(b: Byte; Len: Integer): TBytes; inline;
begin
  SetLength(Result, Len);
  FillChar(Result[0], Len, b);
end;

function XorBlock(s, x: TBytes): TBytes; inline;
var
  L, i: Integer;
  Ps, Px: PByte;
begin
  L := Length(s);
  SetLength(Result, L);
  Ps := PByte(@s[0]);
  Px := PByte(@x[0]);
  for i := 0 to L-1 do
  begin
    Result[i] := Ps^ xor Px^;
    inc(Ps);
    inc(Px);
  end;
end;

function Calc_HMAC_SHA1(message, key: TBytes): TBytes;
const
  blockSize = 64;
begin
  if Length(key) > blocksize then
    key := Calc_SHA1(key);
  key := RPad(key, 0, blocksize);

  Result := Calc_SHA1(ConcatToByteArray(XorBlock(key, Fill($36, blocksize)), message));
  Result := Calc_SHA1(ConcatToByteArray(XorBlock(key, Fill($5c, blockSize)), Result));

  //Result := Calc_SHA1(XorBlock(key, Fill($36, blocksize)) + message);
  //Result := Calc_SHA1(XorBlock(key, Fill($5c, blocksize)) + Result);
end;

// https://keit.co/dcpcrypt-hmac-rfc2104/
function PBKDF2_HMAC_SHA1(pass, salt: TBytes; count, kLen: Integer): TBytes;

  function IntX(i: LongInt): TBytes;
  type
    Int4 = record
      i24, i16, i8, i0: byte;
    end;
  begin
    SetLength(Result, 4);
    Result[0] := Int4(i).i0;
    Result[1] := Int4(i).i8;
    Result[2] := Int4(i).i16;
    Result[3] := Int4(i).i24;
  end;

var
  D, I, J: Integer;
  T, F, U: TBytes;
begin
  T := nil;
  D := Ceil(kLen / 20);  //(hash.GetHashSize div 8));
  for i := 1 to D do
  begin
    F := Calc_HMAC_SHA1(ConcatToByteArray(salt, IntX(i)), pass);
    U := F;
    for j := 2 to count do
    begin
      U := Calc_HMAC_SHA1(U, pass);
      F := XorBlock(F, U);
    end;
    T := ConcatToByteArray(T, F);  // T := T + F;
  end;
  Result := nil;
  SetLength(Result, kLen);
  Move(T[0], Result[0], kLen);
//  Result := Copy(T, 1, kLen);
end;

function Decrypt_AES_ECB(const Key; KeySizeBits: LongWord;
  const InData; var OutData; DataSize: LongWord): String;
{$IFDEF WOLFGANG_EHRHARDT_LIB}
var
  ctx: TAESContext;
  err: Integer;
begin
  err := AES_ECB_Init_Decr(Key, KeySizeBits, ctx{%H-});
  if err <> 0 then
  begin
    Result := 'Decrypt init error ' + IntToStr(err);
    exit;
  end;

  err := AES_ECB_Decrypt(@InData, @OutData, DataSize, ctx);
  if err <> 0 then
  begin
    Result := 'Decrypt error: ' + IntToStr(err);
    exit;
  end;
end;
{$ENDIF}
{$IFDEF DCPCRYPT}
var
  AES_Cipher: TDCP_rijndael;
begin
  Result := '';  // Error message
  AES_Cipher := TDCP_rijndael.Create(nil);
  try
    AES_Cipher.Init(Key, keySizeBits, nil);
    AES_Cipher.DecryptECB(InData, OutData);
  finally
    AES_Cipher.Free;
  end;
end;
{$ENDIF}

function DecryptStream_AES_ECB(const Key; KeySizeBits: LongWord;
  ASrcStream, ADestStream: TStream; ASrcStreamSize: QWord): String;
var
  {$IFDEF WOLFGANG_EHRHARDT_LIB}
  ctx: TAESContext;
  {$ENDIF}
  {$IFDEF DCPCRYPT}
  AES_Cipher: TDCP_rijndael;
  {$ENDIF}
  keySizeBytes: Integer;
  inData: TBytes = nil;
  outData: TBytes = nil;
begin
  Result := '';

  keySizeBytes := KeySizeBits div 8;
  SetLength(inData, keySizeBytes);
  SetLength(outData, keySizeBytes);

  {$IFDEF WOLFGANG_EHRHARDT_LIB}
  AES_ECB_Init_Decr(Key, KeySizeBits, ctx);
  {$ENDIF}

  {$IFDEF DCPCRYPT}
  AES_Cipher := TDCP_rijndael.Create(nil);
  try
    AES_Cipher.Init(Key, KeySizeBits, nil);
  {$ENDIF}

    while ASrcStreamSize > 0 do
    begin
      ASrcStream.ReadBuffer(inData[0], keySizeBytes);
      {$IFDEF WOLFGANG_EHRHARDT_LIB}
      AES_ECB_Decrypt(@inData[0], @outData[0], keySizeBytes, ctx);
      {$ENDIF}
      {$IFDEF DCPCRYPT}
      AES_Cipher.DecryptECB(inData[0], outData[0]);
      {$ENDIF}

      if ASrcStreamSize < keySizeBytes then
        ADestStream.WriteBuffer(outData[0], ASrcStreamSize) // Last block less then key size
      else
        ADestStream.WriteBuffer(outData[0], keySizeBytes);

      if ASrcStreamSize < keySizeBytes then
        ASrcStreamSize := 0
      else
        Dec(ASrcStreamSize, keySizeBytes);
    end;
  {$IFDEF DCPCRYPT}
  finally
    AES_Cipher.Free;
  end;
  {$ENDIF}
end;

{ Decrypts the data in the stream ASrcStream and stores the result in the
  stream ADestStream.
  Decryption algorithm is AES method CBC.
  The hashed password is provided as parameter Key, its length in bits is
  given by KeySizeBits.
  The initialization vector is specified in InitVector.
  If an error occurs, the function result returns an error message with
  error code (error codes are listed in unit AES_Type). Otherwise the
  function result is an empty string. }
function Decrypt_AES_CBC(const Key; KeySizeBits: LongWord; InitVector: Pointer;
  ASrcStream, ADestStream: TStream): String;
{$IFDEF WOLFGANG_EHRHARDT_LIB}
const
  BUF_SIZE = $4000; {must be a multiple of AESBLKSIZE=16 for CBC}
var
  ctx: TAESContext;
  buffer: array[0..BUF_SIZE-1] of byte;
  len, err: Integer;
  n: Word;
begin
  Result := '';

  err := AES_CBC_Init_Decr(Key, KeySizebits, PAESBlock(InitVector)^, ctx{%H-});
  if err <> 0 then
  begin
    Result := 'Decrypt init error ' + IntToStr(err);
    exit;
  end;

  len := ASrcStream.Size;
  while len > 0 do
  begin
    if len > SizeOf(buffer) then
      n := SizeOf(buffer)
    else
      n := len;
    ASrcStream.Read(buffer{%H-}, n);
    dec(len, n);
    err := AES_CBC_Decrypt(@buffer, @buffer, n, ctx);
    if err <> 0 then
    begin
      Result := 'Decrypt error: ' + IntToStr(err);
      exit;
    end;
    ADestStream.Write(buffer, n);
  end;
end;
{$ENDIF}

{$IFDEF DCPCRYPT}
var
  AES_cipher: TDCP_rijndael;
begin
  AES_cipher := TDCP_rijndael.Create(nil);
  try
    AES_cipher.Init(Key, KeySizebits, InitVector);
    AES_cipher.CipherMode := cmCBC;
    AES_cipher.DecryptStream(ASrcStream, ADestStream, ASrcStream.Size);
  finally
    AES_cipher.Free;
  end;
end;
{$ENDIF}

function VerifyDecrypt(AStream: TStream; CheckSum, ChecksumType: string): boolean;
var
  p: Int64;
  buffer: Array of byte = nil;
  expCheckSum: TBytes;
  currCheckSumSHA1: TSHA1Digest;
  lSHA256: TSHA256;
  currChecksumSHA256: TSHA256Digest;
begin
  Result := false;
  p := AStream.Position;
  expCheckSum := DecodeBase64(CheckSum);
  case Uppercase(ChecksumType) of
    'SHA1/1K', 'SHA1-1K':
      begin
        SetLength(buffer, 1024);
        AStream.Write(buffer[0], Length(buffer));
        currCheckSumSHA1 := SHA1Buffer(buffer[0], 1024);
        if Length(expCheckSum) = Length(TSHA1Digest) then
          Result := CompareMem(@expCheckSum[0], @currCheckSumSHA1[0], Length(TSHA1Digest));
      end;
    'SHA256/1K', 'SHA256-1K':
      begin
        SetLength(buffer, 1024);
        AStream.Read(buffer[0], Length(buffer));
        lSHA256.Init;
        lSHA256.Update(@buffer[0], Length(Buffer));
        lSHA256.Final;
        currCheckSumSHA256 := lSHA256.Digest;
        if Length(expCheckSum) = Length(TSHA256Digest) then
          Result := CompareMem(@expChecksum[0], @currCheckSumSHA256[0], Length(TSHA256Digest));
      end;
  end;
  AStream.Position := p;
end;

end.

