unit fpsCrypto;

interface

uses
  SysUtils;

function ExcelPasswordHash(const APassword: String): String;

implementation

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
