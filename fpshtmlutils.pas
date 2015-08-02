unit fpsHTMLUtils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, contnrs;

type
  TsHTMLEntity = record
    E: String;
    Ch: String;
    N: Word;
  end;

function CleanHTMLString(AText: String): String;
function IsHTMLEntity(AText: PChar; out AEntity: TsHTMLEntity): Boolean;

type
  TsHTMLAttr = class
    Name: String;
    Value: String;
    constructor Create(AName, AValue: String);
  end;

  TsHTMLAttrList = class(TObjectList)
  private
    function GetItem(AIndex: Integer): TsHTMLAttr;
    procedure SetItem(AIndex: Integer; AValue: TsHTMLAttr);
  public
    function IndexOfName(AName: String): Integer;
    procedure Parse(AHTML: String);
    property Items[AIndex: Integer]: TsHTMLAttr read GetItem write SetItem; default;
  end;


implementation

uses
  Strings;

const
  // http://unicode.e-workers.de/entities.php
  HTMLEntities: array[0..250] of TsHTMLEntity = (
  // A
    (E: 'Acirc';  Ch: 'Â';  N: 194),   // 0
    (E: 'acirc';  Ch: 'â';  N: 226),
    (E: 'acute';  Ch: '´';  N: 180),
    (E: 'AElig';  Ch: 'Æ';  N: 198),
    (E: 'aelig';  Ch: 'æ';  N: 230),
    (E: 'Agrave'; Ch: 'À';  N: 192),
    (E: 'agrave'; Ch: 'à';  N: 224),
    (E: 'alefsym';Ch: 'ℵ';  N: 8501),
    (E: 'Alpha';  Ch: 'Α';  N: 913),
    (E: 'alpha';  Ch: 'α';  N: 945),
    (E: 'amp';    Ch: '&';  N: 38),    // 10
    (E: 'and';    Ch: '∧';  N: 8743),
    (E: 'ang';    Ch: '∠';  N: 8736),
    (E: 'apos';   Ch: ''''; N: 39),
    (E: 'Aring';  Ch: 'Å';  N: 197),
    (E: 'aring';  Ch: 'å';  N: 229),
    (E: 'asymp';  Ch: '≈';  N: 2248),
    (E: 'Atilde'; Ch: 'Ã';  N: 195),
    (E: 'atilde'; Ch: 'ã';  N: 227),
    (E: 'Auml';   Ch: 'Ä';  N: 196),
    (E: 'auml';   Ch: 'ä';  N: 228),   // 20
  // B
    (E: 'bdquo';  Ch: '„';  N: 8222),  // 21
    (E: 'Beta';   Ch: 'Β';  N: 914),
    (E: 'beta';   Ch: 'β';  N: 946),
    (E: 'brvbar'; Ch: '¦';  N: 166),
    (E: 'bull';   Ch: '•';  N: 8226),
  // C
    (E: 'cap';    Ch: '∩';  N: 8745),  // 26
    (E: 'Ccedil'; Ch: 'Ç';  N: 199),
    (E: 'ccedil'; Ch: 'ç';  N: 231),
    (E: 'cedil';  Ch: '¸';  N: 184),
    (E: 'cent';   Ch: '¢';  N: 162),  // 30
    (E: 'Chi';    Ch: 'Χ';  N: 935),
    (E: 'chi';    Ch: 'χ';  N: 967),
    (E: 'circ';   Ch: 'ˆ';  N: 710),
    (E: 'clubs';  Ch: '♣';  N: 9827),
    (E: 'cong';   Ch: '≅';  N: 8773),  // approximately equal
    (E: 'copy';   Ch: '©';  N: 169),
    (E: 'crarr';  Ch: '↵';  N: 8629),  // carriage return
    (E: 'cup';    Ch: '∪';  N: 8746),
    (E: 'curren'; Ch: '¤';  N: 164),
  // D
    (E: 'Dagger'; Ch: '‡';  N: 8225),  // 40
    (E: 'dagger'; Ch: '†';  N: 8224),
    (E: 'dArr';   Ch: '⇓';  N: 8659),  // wide down-arrow
    (E: 'darr';   Ch: '↓';  N: 8595),  // narrow down-arrow
    (E: 'deg';    Ch: '°';  N: 176),
    (E: 'Delta';  Ch: 'Δ';  N: 916),
    (E: 'delta';  Ch: 'δ';  N: 948),
    (E: 'diams';  Ch: '♦';  N: 9830),
    (E: 'divide'; Ch: '÷';  N: 247),
  // E
    (E: 'Eacute'; Ch: 'É';  N: 201),   // 49
    (E: 'eacute'; Ch: 'é';  N: 233),   // 50
    (E: 'Ecirc';  Ch: 'Ê';  N: 202),
    (E: 'ecirc';  Ch: 'ê';  N: 234),
    (E: 'Egrave'; Ch: 'È';  N: 200),
    (E: 'egrave'; Ch: 'è';  N: 232),
    (E: 'empty';  Ch: '∅';  N: 8709),
    (E: 'emsp';   Ch: ' ';  N: 8195),  // Space character width of "m"
    (E: 'ensp';   Ch: ' ';  N: 8194),  // Space character width of "n"
    (E: 'Epsilon';Ch: 'Ε';  N: 917),   // capital epsilon
    (E: 'epsilon';Ch: 'ε';  N: 949),
    (E: 'equiv';  Ch: '≡';  N: 8801),  // 60
    (E: 'Eta';    Ch: 'Η';  N: 919),
    (E: 'eta';    Ch: 'η';  N: 951),
    (E: 'ETH';    Ch: 'Ð';  N: 208),
    (E: 'eth';    Ch: 'ð';  N: 240),
    (E: 'Euml';   Ch: 'Ë';  N: 203),
    (E: 'euml';   Ch: 'ë';  N: 235),
    (E: 'euro';   Ch: '€';  N: 8364),
    (E: 'exist';  Ch: '∃';  N: 8707),
  // F
    (E: 'fnof';   Ch: 'ƒ';  N: 402),   // 70
    (E: 'forall'; Ch: '∀';  N: 8704),
    (E: 'frac12'; Ch: '½';  N: 189),
    (E: 'frac14'; Ch: '¼';  N: 188),
    (E: 'frac34'; Ch: '¾';  N: 190),
    (E: 'frasl';  Ch: '⁄';  N: 8260),
  // G
    (E: 'Gamma';  Ch: 'Γ';  N: 915),
    (E: 'gamma';  Ch: 'γ';  N: 947),
    (E: 'ge';     Ch: '≥';  N: 8805),
    (E: 'gt';     Ch: '>';  N: 62),
  // H
    (E: 'hArr';   Ch: '⇔';  N: 8660),  // 80, wide horizontal double arrow
    (E: 'harr';   Ch: '↔';  N: 8596),  // narrow horizontal double arrow
    (E: 'hearts'; Ch: '♥';  N: 9829),
    (E: 'hellip'; Ch: '…';  N: 8230),
  // I
    (E: 'Iacute'; Ch: 'Í';  N: 205),
    (E: 'iacute'; Ch: 'í';  N: 237),
    (E: 'Icirc';  Ch: 'Î';  N: 206),
    (E: 'icirc';  Ch: 'î';  N: 238),
    (E: 'iexcl';  Ch: '¡';  N: 161),
    (E: 'Igrave'; Ch: 'Ì';  N: 204),
    (E: 'igrave'; Ch: 'ì';  N: 236),   // 90
    (E: 'image';  Ch: 'ℑ';  N: 2465),  // I in factura
    (E: 'infin';  Ch: '∞';  N: 8734),
    (E: 'int';    Ch: '∫';  N: 8747),
    (E: 'Iota';   Ch: 'Ι';  N: 921),
    (E: 'iota';   Ch: 'ι';  N: 953),
    (E: 'iquest'; Ch: '¿';  N: 191),
    (E: 'isin';   Ch: '∈';  N: 8712),
    (E: 'Iuml';   Ch: 'Ï';  N: 207),
    (E: 'iuml';   Ch: 'ï';  N: 239),
  // K
    (E: 'Kappa';  Ch: 'Κ';  N: 922),  // 100
    (E: 'kappa';  Ch: 'κ';  N: 254),
  // L
    (E: 'Lambda'; Ch: 'Λ';  N: 923),
    (E: 'lambda'; Ch: 'λ';  N: 955),
    (E: 'lang';   Ch: '⟨';  N: 9001),  // Left-pointing angle bracket
    (E: 'laquo';  Ch: '«';  N: 171),
    (E: 'lArr';   Ch: '⇐';  N: 8656),  // Left-pointing wide arrow
    (E: 'larr';   Ch: '←';  N: 8592),
    (E: 'lceil';  Ch: '⌈';  N: 8968),  // Left ceiling
    (E: 'ldquo';  Ch: '“';  N: 8220),
    (E: 'le';     Ch: '≤';  N: 8804),  // 110
    (E: 'lfloor'; Ch: '⌊';  N: 8970),  // Left floor
    (E: 'lowast'; Ch: '∗';  N: 8727),  // Low asterisk
    (E: 'loz';    Ch: '◊';  N: 9674),
    (E: 'lrm';    Ch: '‎';  N: 8206),  // Left-to-right mark
    (E: 'lsaquo'; Ch: '‹';  N: 8249),
    (E: 'lsquo';  Ch: '‘';  N: 8216),
    (E: 'lt';     Ch: '<';  N: 60),
  // M
    (E: 'macr';   Ch: '¯';  N: 175),
    (E: 'mdash';  Ch: '—';  N: 8212),
    (E: 'micro';  Ch: 'µ';  N: 181),  // 120
    (E: 'middot'; Ch: '·';  N: 183),
    (E: 'minus';  Ch: '−';  N: 8722),
    (E: 'Mu';     Ch: 'Μ';  N: 924),
    (E: 'mu';     Ch: 'μ';  N: 956),
  // N
    (E: 'nabla';  Ch: '∇';  N: 8711),
    (E: 'nbsp';   Ch: ' ';  N: 160),   // 126
    (E: 'ndash';  Ch: '–';  N: 8211),
    (E: 'ne';     Ch: '≠';  N: 8800),
    (E: 'ni';     Ch: '∋';  N: 8715),
    (E: 'not';    Ch: '¬';  N: 172),   // 130
    (E: 'notin';  Ch: '∉';  N: 8713),  // math: "not in"
    (E: 'nsub';   Ch: '⊄';  N: 8836),  // math: "not a subset of"
    (E: 'Ntilde'; Ch: 'Ñ';  N: 209),
    (E: 'ntilde'; Ch: 'ñ';  N: 241),
    (E: 'Nu';     Ch: 'Ν';  N: 925),
    (E: 'nu';     Ch: 'ν';  N: 957),
  // O
    (E: 'Oacute'; Ch: 'Ó';  N: 211),
    (E: 'oacute'; Ch: 'ó';  N: 243),
    (E: 'Ocirc';  Ch: 'Ô';  N: 212),
    (E: 'ocirc';  Ch: 'ô';  N: 244),
    (E: 'OElig';  Ch: 'Œ';  N: 338),
    (E: 'oelig';  Ch: 'œ';  N: 339),
    (E: 'Ograve'; Ch: 'Ò';  N: 210),
    (E: 'ograve'; Ch: 'ò';  N: 242),
    (E: 'oline';  Ch: '‾';  N: 8254),
    (E: 'Omega';  Ch: 'Ω';  N: 937),
    (E: 'omega';  Ch: 'ω';  N: 969),
    (E: 'Omicron';Ch: 'Ο';  N: 927),
    (E: 'omicron';Ch: 'ο';  N: 959),
    (E: 'oplus';  Ch: '⊕';  N: 8853),  // Circled plus
    (E: 'or';     Ch: '∨';  N: 8744),
    (E: 'ordf';   Ch: 'ª';  N: 170),
    (E: 'ordm';   Ch: 'º';  N: 186),
    (E: 'Oslash'; Ch: 'Ø';  N: 216),
    (E: 'oslash'; Ch: 'ø';  N: 248),
    (E: 'Otilde'; Ch: 'Õ';  N: 213),
    (E: 'otilde'; Ch: 'õ';  N: 245),
    (E: 'otimes'; Ch: '⊗';  N: 8855),  // Circled times
    (E: 'Ouml';   Ch: 'Ö';  N: 214),
    (E: 'ouml';   Ch: 'ö';  N: 246),
  // P
    (E: 'para';   Ch: '¶';  N: 182),
    (E: 'part';   Ch: '∂';  N: 8706),
    (E: 'permil'; Ch: '‰';  N: 8240),
    (E: 'perp';   Ch: '⊥';  N: 8869),
    (E: 'Phi';    Ch: 'Φ';  N: 934),
    (E: 'phi';    Ch: 'φ';  N: 966),
    (E: 'Pi';     Ch: 'Π';  N: 928),
    (E: 'pi';     Ch: 'π';  N: 960),  // lower-case pi
    (E: 'piv';    Ch: 'ϖ';  N: 982),
    (E: 'plusmn'; Ch: '±';  N: 177),
    (E: 'pound';  Ch: '£';  N: 163),
    (E: 'Prime';  Ch: '″';  N: 8243),
    (E: 'prime';  Ch: '′';  N: 8242),
    (E: 'prod';   Ch: '∏';  N: 8719),
    (E: 'prop';   Ch: '∝';  N: 8733),
    (E: 'Psi';    Ch: 'Ψ';  N: 936),
    (E: 'psi';    Ch: 'ψ';  N: 968),
  // Q
    (E: 'quot';   Ch: '"';  N: 34),
  // R
    (E: 'radic';  Ch: '√';  N: 8730),
    (E: 'rang';   Ch: '⟩';  N: 9002),  // right-pointing angle bracket
    (E: 'raquo';  Ch: '»';  N: 187),
    (E: 'rArr';   Ch: '⇒';  N: 8658),
    (E: 'rarr';   Ch: '→';  N: 8594),
    (E: 'rceil';  Ch: '⌉';  N: 8969),  // right ceiling
    (E: 'rdquo';  Ch: '”';  N: 8221),
    (E: 'real';   Ch: 'ℜ';  N: 8476),  // R in factura
    (E: 'reg';    Ch: '®';  N: 174),
    (E: 'rfloor'; Ch: '⌋';  N: 8971),  // Right floor
    (E: 'Rho';    Ch: 'Ρ';  N: 929),
    (E: 'rho';    Ch: 'ρ';  N: 961),
    (E: 'rlm';    Ch: '';   N: 8207),   // right-to-left mark
    (E: 'rsaquo'; Ch: '›';  N: 8250),
    (E: 'rsquo';  Ch: '’';  N: 8217),

  // S
    (E: 'sbquo';  Ch: '‚';  N: 8218),
    (E: 'Scaron'; Ch: 'Š';  N: 352),
    (E: 'scaron'; Ch: 'š';  N: 353),
    (E: 'sdot';   Ch: '⋅';  N: 8901),  // math: dot operator
    (E: 'sect';   Ch: '§';  N: 167),
    (E: 'shy';    Ch: '';   N: 173),   // conditional hyphen
    (E: 'Sigma';  Ch: 'Σ';  N: 931),
    (E: 'sigma';  Ch: 'σ';  N: 963),
    (E: 'sigmaf'; Ch: 'ς';  N: 962),
    (E: 'sim';    Ch: '∼';  N: 8764),  // similar
    (E: 'spades'; Ch: '♠';  N: 9824),
    (E: 'sub';    Ch: '⊂';  N: 8834),
    (E: 'sube';   Ch: '⊆';  N: 8838),
    (E: 'sum';    Ch: '∑';  N: 8721),
    (E: 'sup';    Ch: '⊃';  N: 8835),
    (E: 'sup1';   Ch: '¹';  N: 185),
    (E: 'sup2';   Ch: '²';  N: 178),
    (E: 'sup3';   Ch: '³';  N: 179),
    (E: 'supe';   Ch: '⊇';  N: 8839),
    (E: 'szlig';  Ch: 'ß';  N: 223),
  //T
    (E: 'Tau';    Ch: 'Τ';  N: 932),
    (E: 'tau';    Ch: 'τ';  N: 964),
    (E: 'there4'; Ch: '∴';  N: 8756),
    (E: 'Theta';  Ch: 'Θ';  N: 920),
    (E: 'theta';  Ch: 'θ';  N: 952),
    (E: 'thetasym';Ch: 'ϑ'; N: 977),
    (E: 'thinsp'; Ch: ' ';  N: 8201),  // thin space
    (E: 'THORN';  Ch: 'Þ';  N: 222),
    (E: 'thorn';  Ch: 'þ';  N: 254),
    (E: 'tilde';  Ch: '˜';  N: 732),
    (E: 'times';  Ch: '×';  N: 215),
    (E: 'trade';  Ch: '™';  N: 8482),
  // U
    (E: 'Uacute'; Ch: 'Ú';  N: 218),
    (E: 'uacute'; Ch: 'ú';  N: 250),
    (E: 'uArr';   Ch: '⇑';  N: 8657),  // wide up-arrow
    (E: 'uarr';   Ch: '↑';  N: 8593),
    (E: 'Ucirc';  Ch: 'Û';  N: 219),
    (E: 'ucirc';  Ch: 'û';  N: 251),
    (E: 'Ugrave'; Ch: 'Ù';  N: 217),
    (E: 'ugrave'; Ch: 'ù';  N: 249),
    (E: 'uml';    Ch: '¨';  N: 168),
    (E: 'upsih';  Ch: 'ϒ';  N: 978),
    (E: 'Upsilon';Ch: 'Υ';  N: 933),
    (E: 'upsilon';Ch: 'υ';  N: 965),
    (E: 'Uuml';   Ch: 'Ü';  N: 220),
    (E: 'uuml';   Ch: 'ü';  N: 252),
  // W
    (E: 'weierp'; Ch: '℘';  N: 8472),  // Script Capital P; Weierstrass Elliptic Function
  // X
    (E: 'Xi';     Ch: 'Ξ';  N: 926),
    (E: 'xi';     Ch: 'ξ';  N: 958),
  // Y
    (E: 'Yacute'; Ch: 'Ý';  N: 221),
    (E: 'yacute'; Ch: 'ý';  N: 253),
    (E: 'yen';    Ch: '¥';  N: 165),
    (E: 'Yuml';   Ch: 'Ÿ';  N: 376),
    (E: 'yuml';   Ch: 'ÿ';  N: 255),
  // Z
    (E: 'Zeta';   Ch: 'Ζ';  N: 918),
    (E: 'zeta';   Ch: 'ζ';  N: 950),
    (E: 'zwj';    Ch: '';   N: 8205),   // Zero-width joiner
    (E: 'zwnj';   Ch: '';   N: 8204)    // Zero-width non-joiner
  );

function IsHTMLEntity(AText: PChar; out AEntity: TsHTMLEntity): Boolean;

  function Compare(s: String): Boolean;
  var
    j: Integer;
  begin
    Result := false;
    for j:=1 to Length(s) do
      if s[j] <> PChar(AText)[j-1] then
        exit;
    if PChar(AText)[Length(s)] <> ';' then
      exit;
    Result := true;
  end;

var
  j, k: Integer;
  equ: Boolean;
  ch1, ch2: Char;
  P: PChar;

begin
  Result := false;
  for k:=0 to High(HTMLEntities) do
  begin
    equ := Compare(HTMLEntities[k].E);
    if not equ then
    begin
      P := AText;
      ch1 := P^;
      if ch1 = '#' then
      begin
        inc(P);
        ch2 := P^;
        if ch1 = 'x' then
          equ := Compare(Format('#x%x', [HTMLEntities[k].N]))
        else
          equ := Compare(Format('#%d', [HTMLEntities[k].N]));
      end;
    end;
    if equ then
    begin
      AEntity := HTMLEntities[k];
      Result := true;
      exit;
    end;
  end;
end;

function CleanHTMLString(AText: String): String;
var
  len: Integer;
  ent: TsHTMLEntity;
  P: PChar;
  ch: Char;
begin
  Result := '';

  // Remove leading and trailing spaces and line endings coming from formatted
  // source lines
  while (Length(AText) > 0) and (AText[1] in [#9, #10, #13, ' ']) do
    Delete(AText, 1,1);
  while (Length(AText) > 0) and (AText[Length(AText)] in [#9, #10, #13, ' ']) do
    Delete(AText, Length(AText), 1);
  if AText = '' then
    exit;

  // Replace HTML entities by their counter part UTF8 characters
  len := Length(AText);
  P := @AText[1];
  while (P^ <> #0) do begin
    ch := P^;
    case ch of
      '&': begin
             inc(P);
             if (P <> nil) and IsHTMLEntity(P, ent) then
             begin
               Result := Result + ent.Ch;
               inc(P, Length(ent.E));
             end else
             begin
               Result := Result + '&';
               Continue;
             end;
           end;
      else Result := Result + ch;
    end;
    inc(P);
  end;
end;


{==============================================================================}
{                                TsHTMLAttr                                    }
{==============================================================================}

constructor TsHTMLAttr.Create(AName, AValue: String);
begin
  Name := AName;
  Value := AValue;
end;


{==============================================================================}
{                              TsHTMLAttrList                                  }
{==============================================================================}

function TsHTMLAttrList.GetItem(AIndex: Integer): TsHTMLAttr;
begin
  Result := TsHTMLAttr(inherited GetItem(AIndex));
end;

function TsHTMLAttrList.IndexOfName(AName: String): Integer;
begin
  AName := Lowercase(AName);
  for Result := 0 to Count-1 do
    if GetItem(Result).Name = AName then
      exit;
  Result := -1;
end;

{ AHTML is a HTML string beginning with a < tag. Seeks the first space to split
  off the HTML tag. Then seeks for = and " characters to extract the attributes
  which are split into name/value pairs at the = character. The value part is
  unquoted. }
procedure TsHTMLAttrList.Parse(AHTML: String);
var
  i: Integer;
  len: Integer;
  value, nam: String;
begin
  Clear;
  if (AHTML[1] <> '<') then           // just for simplification
    raise Exception.Create('[THTMLAttrList.Parse] HTML tags expected.');

  // Find first space
  i := 1;
  len := Length(AHTML);
  while (i <= len) and (AHTML[i] <> ' ') do inc(i);

  // Parse attribute string
  nam := '';
  while (i <= len) do
  begin
    case AHTML[i] of
      '=': begin
             inc(i);
             if AHTML[i] <> '"' then
               raise Exception.Create('[THTMLAttrList.Parse] Quotation marks expected.');
             value := '';
             inc(i);   // skip the initial '"'
             while (AHTML[i] <> '"') do
             begin
               value := value + AHTML[i];
               inc(i);
             end;
             inc(i);  // skip the final '"'
             Add(TsHTMLAttr.Create(lowercase(nam), value));
             nam := '';
           end;
      ' ', '/', '>': ;
      else nam := nam + AHTML[i];
    end;
    inc(i);
  end;
end;

procedure TsHTMLAttrList.SetItem(AIndex: Integer; AValue: TsHTMLAttr);
begin
  inherited SetItem(AIndex, AValue);
end;


end.

