unit fpsCrypto;

interface

uses
  SysUtils, fpsTypes;

type
  TsAlgorithmUsage = (auExcel, auOpenDocument);

function AlgorithmToStr(Algorithm: TsCryptoAlgorithm; AUsage: TsAlgorithmUsage): String;
function StrToAlgorithm(const AName: String): TsCryptoAlgorithm;

function ExcelPasswordHash(const APassword: String): String;

implementation

uses
  LazUTF8;

function AlgorithmToStr(Algorithm: TsCryptoAlgorithm; AUsage: TsAlgorithmUsage): String;
begin
  Result := '';
  case AUsage of
    auExcel:
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
      end;
    auOpenDocument:
      case Algorithm of
        caSHA1       : Result := 'http://www.w3.org/2000/09/xmldsig#sha1';
        caSHA256     : Result := 'http://www.w3.org/2000/09/xmldsig#sha256';
      end;
  end;
end;

function StrToAlgorithm(const AName: String): TsCryptoAlgorithm;
begin
  case AName of
    // Excel
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
    else
      // Libre/OpenOffice
      if pos('sha1', AName) > 0 then    // http://www.w3.org/2000/09/xmldsig#sha1
        Result := caSHA1
      else
      if pos('sha256', AName) > 0 then  // http://www.w3.org/2000/09/xmldsig#sha256
        Result := caSHA256
      else
        Result := caUnknown;
  end;
end;

{@@ This is the code for generating Excel 2010 and earlier password's hash
  See: http://forum.lazarus.freepascal.org/index.php/topic,36075.msg240132.html#msg240132 }
function ExcelPasswordHash( const APassword: string ): string;
var
  i: Integer;
  PassLen: Integer;
  Password: string;
  PassHash: Word = 0;
begin
  // we are needed to work with single byte character.
  Password:= UTF8ToWinCP(APassword);
  PassLen := Length(Password);

  if PassLen = 0 then
  begin
    Result := '';
    exit;
  end;

  for i:= PassLen downto 1 do
  begin
    PassHash:= ((PassHash shr 14) and  $0001) or ((PassHash shl  1) and  $7fff);
    PassHash:= PassHash xor ord(Password[i]);
  end;

  PassHash:= ((PassHash shr 14) and  $0001) or ((PassHash shl  1) and  $7fff);
  PassHash:= PassHash xor PassLen xor $CE4B;

  Result := IntToHex(PassHash, 4);
end;

end.
