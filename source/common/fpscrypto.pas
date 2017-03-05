unit fpsCrypto;

interface

uses
  SysUtils, fpsTypes;

function AlgorithmToStr(Algorithm: TsCryptoAlgorithm): String;
function StrToAlgorithm(const AName: String): TsCryptoAlgorithm;

function ExcelPasswordHash(const APassword: String): String;

implementation

function AlgorithmToStr(Algorithm: TsCryptoAlgorithm): String;
begin
  case Algorithm of
    caExcel      : Result := 'EXCEL';
    caMD2        : Result := 'MD2';
    caMD4        : Result := 'MD4';
    caMD5        : Result := 'MD5';
    caRIPEMD128  : Result := 'RIPEMD-128';
    caRIPEMD160  : Result := 'RIPEMD-160';
    caSHA1       : Result := 'SHA-1';
    caSHA256     : Result := 'SHA-256';
    caSHA384     : Result := 'SHA-384';
    caSHA512     : Result := 'SHA-512';
    caWHIRLPOOL  : Result := 'WHIRLPOOL';
    else           Result := '';
  end;
end;

function StrToAlgorithm(const AName: String): TsCryptoAlgorithm;
begin
  case AName of
    'MD2'        : Result := caMD2;
    'MD4'        : Result := caMD4;
    'MD5'        : Result := caMD5;
    'RIPEMD-128' : Result := caRIPEMD128;
    'RIPEMD-160' : Result := caRIPEMD160;
    'SHA-1'      : Result := caSHA1;
    'SHA-256'    : Result := caSHA256;
    'SHA-384'    : Result := caSHA384;
    'SHA-512'    : Result := caSHA512;
    'WHIRLPOOL'  : Result := caWHIRLPOOL;
    else           Result := caUnknown;
  end;
end;

{@@ This is the code for generating Excel 2010 and earlier password's hash }
function ExcelPasswordHash(const APassword: string): string;
const
  Key = $CE4B;
var
  i: Integer;
  HashValue: Word = 0;
begin
  for i:= Length(APassword) downto 1 do
  begin
    HashValue := ord(APassword[i]) xor HashValue;
    HashValue := HashValue shl 1;
  end;
  HashValue := HashValue xor Length(APassword) xor Key;
 
  Result := IntToHex(HashValue, 4);
end;

end.
