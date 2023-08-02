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
  DCPrijndael;
  {$ENDIF}

function Calc_SHA1(const AText: RawByteString): RawByteString;
function Calc_SHA256(AText: RawByteString): RawByteString;
function PBKDF2_HMAC_SHA1(pass, salt: RawByteString; count, kLen: Integer): RawByteString;

function Decrypt_AES_ECB(const Key; KeySizeBits: LongWord;
  const InData; var OutData; DataSize: LongWord): String;

function DecryptStream_AES_ECB(const Key; KeySizeBits: LongWord;
  ASrcStream, ADestStream: TStream; ASrcStreamSize: QWord): String;


implementation

uses
  Math;

function Calc_SHA1(const AText: RawByteString): RawByteString;
var
  sha1Digest: TSHA1Digest;
begin
  sha1Digest := SHA1String(AText);
  SetLength(Result, 20);
  Move(sha1Digest[0], Result[1], 20);
end;

function Calc_SHA256(AText: RawByteString): RawByteString;
var
  sha256: TSHA256;
begin
  sha256.Init;
  sha256.Update(@AText[1], Length(AText));
  sha256.Final;
  SetLength(Result, 32);
  Move(sha256.Digest[0], Result[1], 32);
end;

function RPad(x: RawByteString; c: Char; s: Integer): RawByteString;
var
  L: Integer;
begin
  L := Length(x);
  if L < s then
  begin
    SetLength(Result, s);
    Move(x[1], Result[1], L);
    FillChar(Result[L+1], s-L, c);
  end else
    Result := x;
end;

function Fill(c: Char; Len: Integer): RawByteString; inline;
begin
  SetLength(Result, Len);
  FillChar(Result[1], Len, c);
end;

function XorBlock(s, x: RawByteString): RawByteString; inline;
var
  L, i: Integer;
  Ps, Px: PByte;
begin
  L := Length(s);
  SetLength(Result, L);
  Ps := PByte(@s[1]);
  Px := PByte(@x[1]);
  for i := 1 to L do
  begin
    Result[i] := Char(Ps^ xor Px^);
    inc(Ps);
    inc(Px);
  end;
end;

function Calc_HMAC_SHA1(message, key: RawByteString): RawByteString;
const
  blockSize = 64;
begin
  if Length(key) > blocksize then
    key := Calc_SHA1(key);
  key := RPad(key, #0, blocksize);

  Result := Calc_SHA1(XorBlock(key, Fill(#$36, blocksize)) + message);
  Result := Calc_SHA1(XorBlock(key, Fill(#$5c, blocksize)) + Result);
end;

// https://keit.co/dcpcrypt-hmac-rfc2104/
function PBKDF2_HMAC_SHA1(pass, salt: RawByteString; count, kLen: Integer): RawByteString;

  function IntX(i: Integer): RawByteString;
  type
    Int4 = record
      i24, i16, i8, i0: char;
    end;
  begin
    SetLength(Result, 4);
    Result[1] := Int4(i).i0;
    Result[2] := Int4(i).i8;
    Result[3] := Int4(i).i16;
    Result[4] := Int4(i).i24;
  end;

var
  D, I, J: Integer;
  T, F, U: RawByteString;
begin
  T := '';
  D := Ceil(kLen / 20);  //(hash.GetHashSize div 8));
  for i := 1 to D do
  begin
    F := Calc_HMAC_SHA1(salt + IntX(i), pass);
    U := F;
    for j := 2 to count do
    begin
      U := Calc_HMAC_SHA1(U, pass);
      F := XorBlock(F, U);
    end;
    T := T + F;
  end;
  Result := Copy(T, 1, kLen);
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

end.

