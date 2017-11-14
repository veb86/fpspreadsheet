unit fpsHTMLUtils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, contnrs, fpstypes, fpspreadsheet;

type
  TsHTMLEntity = record
    E: String;
    Ch: String;
    N: Word;
  end;

function CleanHTMLString(AText: String): String;
function IsHTMLEntity(AText: PChar; out AEntity: TsHTMLEntity): Boolean;
function RemoveHTMLEntities(const AText: String): String;

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
  protected
    procedure ParseStyle(AStyle: String);
  public
    function IndexOfName(AName: String): Integer;
    procedure Parse(AHTML: String);
    property Items[AIndex: Integer]: TsHTMLAttr read GetItem write SetItem; default;
  end;

  TsTagCase = (tcLowercase, tcUppercase, tcProperCase);

procedure HTMLToRichText(AWorkbook: TsWorkbook; AFont: TsFont;
  const AHTMLText: String; out APlainText: String;
  out ARichTextParams: TsRichTextParams);

procedure RichTextToHTML(AWorkbook: TsWorkbook; AFont: TsFont;
  const APlainText: String; const ARichTextParams: TsRichTextParams;
  out AHTMLText: String; APrefix:String = ''; ATagCase: TsTagCase = tcLowercase);


implementation

uses
  math, lazUtf8, fasthtmlparser,
  fpsUtils, fpsClasses;

const
  // http://unicode.e-workers.de/entities.php
  // http://www.utf8-zeichentabelle.de/unicode-utf8-table.pl
  HTMLEntities: array[0..250] of TsHTMLEntity = (
  // A
    (E: 'Acirc';  Ch: 'Â';  N: 194),   // 0
    (E: 'acirc';  Ch: 'â';  N: 226),
    (E: 'acute';  Ch: '´';  N: 180),
    (E: 'AElig';  Ch: 'Æ';  N: 198),
    (E: 'aelig';  Ch: 'æ';  N: 230),
    (E: 'Agrave'; Ch: 'À';  N: 192),
    (E: 'agrave'; Ch: 'à';  N: 224),
    (E: 'alefsym';Ch: #$E2#$84#$B5;  N: 8501),
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
    (E: 'cong';   Ch: #$E2#$89#$85;  N: 8773),  // approximately equal
    (E: 'copy';   Ch: '©';  N: 169),
    (E: 'crarr';  Ch: #$E2#$86#$B5;  N: 8629),  // carriage return
    (E: 'cup';    Ch: '∪';  N: 8746),
    (E: 'curren'; Ch: '¤';  N: 164),
  // D
    (E: 'Dagger'; Ch: '‡';  N: 8225),  // 40
    (E: 'dagger'; Ch: '†';  N: 8224),
    (E: 'dArr';   Ch: #$E2#$87#$93;  N: 8659),  // wide down-arrow
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
    (E: 'empty';  Ch: #$e2#$88#$85;  N: 8709),
    (E: 'emsp';   Ch: #$E2#$80#$83;  N: 8195),  // Space character width of "m"
    (E: 'ensp';   Ch: #$E2#$80#$82;  N: 8194),  // Space character width of "n"
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
    (E: 'image';  Ch: #$E2#$84#$91;  N: 2465),  // I in factura
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
    (E: 'lArr';   Ch: #$E2#$87#$90;  N: 8656),  // Left-pointing wide arrow
    (E: 'larr';   Ch: '←';  N: 8592),
    (E: 'lceil';  Ch: '⌈';  N: 8968),  // Left ceiling
    (E: 'ldquo';  Ch: '“';  N: 8220),
    (E: 'le';     Ch: '≤';  N: 8804),  // 110
    (E: 'lfloor'; Ch: '⌊';  N: 8970),  // Left floor
    (E: 'lowast'; Ch: #$e2#$88#$97;  N: 8727),  // Low asterisk
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
    (E: 'notin';  Ch: #$e2#$88#$89;  N: 8713),  // math: "not in"
    (E: 'nsub';   Ch: #$e2#$8a#$84;  N: 8836),  // math: "not a subset of"
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
    (E: 'oplus';  Ch: #$e2#$8a#$95;  N: 8853),  // Circled plus
    (E: 'or';     Ch: '∨';  N: 8744),
    (E: 'ordf';   Ch: 'ª';  N: 170),
    (E: 'ordm';   Ch: 'º';  N: 186),
    (E: 'Oslash'; Ch: 'Ø';  N: 216),
    (E: 'oslash'; Ch: 'ø';  N: 248),
    (E: 'Otilde'; Ch: 'Õ';  N: 213),
    (E: 'otilde'; Ch: 'õ';  N: 245),
    (E: 'otimes'; Ch: #$E2#$8A#$97;  N: 8855),  // Circled times
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
    (E: 'real';   Ch: #$E2#$84#$9C;  N: 8476),  // R in factura
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
    (E: 'sdot';   Ch: #$E2#$8B#$85;  N: 8901),  // math: dot operator
    (E: 'sect';   Ch: '§';  N: 167),
    (E: 'shy';    Ch: #$C2#$AD; N: 173),   // conditional hyphen
    (E: 'Sigma';  Ch: 'Σ';  N: 931),
    (E: 'sigma';  Ch: 'σ';  N: 963),
    (E: 'sigmaf'; Ch: 'ς';  N: 962),
    (E: 'sim';    Ch: #$E2#$88#$BC;  N: 8764),  // similar
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
    (E: 'thinsp'; Ch: #$E2#$80#$89;  N: 8201),  // thin space
    (E: 'THORN';  Ch: 'Þ';  N: 222),
    (E: 'thorn';  Ch: 'þ';  N: 254),
    (E: 'tilde';  Ch: '˜';  N: 732),
    (E: 'times';  Ch: '×';  N: 215),
    (E: 'trade';  Ch: '™';  N: 8482),
  // U
    (E: 'Uacute'; Ch: 'Ú';  N: 218),
    (E: 'uacute'; Ch: 'ú';  N: 250),
    (E: 'uArr';   Ch: #$E2#$87#$91;  N: 8657),  // wide up-arrow
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
    (E: 'weierp'; Ch: #$E2#$84#$98;  N: 8472),  // Script Capital P; Weierstrass Elliptic Function
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
  k: Integer;
  equ: Boolean;
  ch1: Char;
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
  ent: TsHTMLEntity;
  P: PChar;
  ch: Char;
  hasStartSpace, hasEndSpace: Boolean;
begin
  Result := '';

  // Remove leading and trailing spaces and line endings coming from formatted
  // source lines. Retain 1 single space, at the end even without spaces found.
  // No idea if this is 100% correct - at least, looks good.
  hasStartSpace := false;
  while (Length(AText) > 0) and (AText[1] in [#9, #13, #10, ' ']) do
  begin
    if AText[1] = ' ' then hasStartSpace := true; // A leading space will be added later
    Delete(AText, 1, 1);
  end;

  hasEndSpace := false;
  while (Length(AText) > 0) and (AText[Length(AText)] in [#9, #10, #13, ' ']) do
  begin
    hasEndSpace := true;                 // A trailing space will be added later
    Delete(AText, Length(AText), 1);
  end;

  if AText = '' then
  begin
    if hasStartSpace or hasEndSpace then Result := ' ';
    exit;
  end;

  // Replace HTML entities by their counter part UTF8 characters
  P := @AText[1];
  while (P^ <> #0) do begin
    ch := P^;
    case ch of
      ' ': begin
             // collapse multiple spaces to a single space (HTML spec)
             // http://stackoverflow.com/questions/24615355/browser-white-space-rendering
             Result := Result + ' ';
             inc(P);
             while (P^ = ' ') do inc(P);
             dec(P);
           end;
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
      else
           Result := Result + ch;
    end;
    inc(P);
  end;

  // Add leading and trailing spaces from above.
  if hasStartSpace then Result := ' ' + Result;
  if hasEndSpace then Result := Result + ' ';
end;

function RemoveHTMLEntities(const AText: String): String;
var
  ent: TsHTMLEntity;
  P: PChar;
  ch: AnsiChar;
begin
  Result := '';
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
    raise EFPSpreadsheet.Create('[THTMLAttrList.Parse] HTML tags expected.');

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
             value := '';
             if AHTML[i] = '"' then
             begin
               inc(i);  // skip the initial '"'
               while AHTML[i] <> '"' do
               begin
                 value := value + AHTML[i];
                 inc(i);
               end;
               inc(i);  // skip the final '"'
             end else
               while not (AHTML[i] in [' ', '>', '/']) do
               begin
                 value := value + AHTML[i];
                 inc(i);
               end;
             Add(TsHTMLAttr.Create(lowercase(trim(nam)), trim(value)));
             nam := '';
           end;
      ' ', '/', '>': ;
      else nam := nam + AHTML[i];
    end;
    inc(i);
  end;

  i := IndexOfName('style');
  if i > -1 then ParseStyle(Items[i].Value);
end;

{ AStyle is the value part of a 'style="...."' HTML string. Splits the into
  individual records at the semicolons (;) and into name-value pairs at the
  colon (:). Adds the name-value pairs to the list. }
procedure TsHTMLAttrList.ParseStyle(AStyle: String);
var
  i, len: Integer;
  value, nam: String;
begin
  i := 1;
  len := Length(AStyle);

  // skip white space
  while (i <= len) and (AStyle[i] = ' ') do inc(i);

  // iterate through string
  nam := '';
  while (i <= len) do
  begin
    case AStyle[i] of
      ':': begin  // name-value separator
             inc(i);  // skip ':' ...
             while (i <= len) and (AStyle[i] = ' ') do inc(i);  // ... and white space
             value := '';
             while (i <= len) and (AStyle[i] <> ';') do
             begin
               value := value + AStyle[i];
               inc(i);
             end;
             Add(TsHTMLAttr.Create(lowercase(trim(nam)), UnquoteStr(trim(value))));
             nam := '';
           end;
      ' ': ;   // skip white space
      else nam := nam + AStyle[i];
    end;
    inc(i);
  end;
end;

procedure TsHTMLAttrList.SetItem(AIndex: Integer; AValue: TsHTMLAttr);
begin
  inherited SetItem(AIndex, AValue);
end;


{==============================================================================}
{                         HTML-to-Rich-text conversion                         }
{==============================================================================}
type
  TsHTMLAnalyzer = class(THTMLParser)
  private
    FWorkbook: TsWorkbook;
    FPlainText: String;
    FRichTextParams: TsRichTextParams;
    FAttrList: TsHTMLAttrList;
    FFontStack: TsIntegerStack;
    FCurrFont: TsFont;
    FPointSeparatorSettings: TFormatSettings;
    FPreserveSpaces: Boolean;
    function AddFont(AFont: TsFont): Integer;
    procedure AddRichTextParam(AFont: TsFont; AHyperlinkIndex: Integer = -1);
    procedure ProcessFontRestore;
    procedure ReadFont(AFont: TsFont);
    procedure TagFoundHandler(NoCaseTag, ActualTag: string);
    procedure TextFoundHandler(AText: string);
  public
    constructor Create(AWorkbook: TsWorkbook; AFont: TsFont; AText: String);
    destructor Destroy; override;
    property PlainText: String read FPlainText;
    property RichTextParams: TsRichTextParams read FRichTextParams;
    property PreserveSpaces: Boolean read FPreserveSpaces write FPreserveSpaces;
  end;

constructor TsHTMLAnalyzer.Create(AWorkbook: TsWorkbook; AFont: TsFont;
  AText: String);
begin
  if AWorkbook = nil then
    raise EFPSpreadsheet.Create('[TsHTMLAnalyzer.Create] Workbook required.');
  if AFont = nil then
    raise EFPSpreadsheet.Create('[TsHTMLAnalyzer.Create] Font required.');

  inherited Create(AText);
  FWorkbook := AWorkbook;

  OnFoundTag := @TagFoundHandler;
  OnFoundText := @TextFoundHandler;

  FPlainText := '';
  SetLength(FRichTextParams, 0);

  FAttrList := TsHTMLAttrList.Create;
  FCurrFont := TsFont.Create;
  FCurrFont.CopyOf(AFont);

  FFontStack := TsIntegerStack.Create;

  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';
end;

destructor TsHTMLAnalyzer.Destroy;
begin
  FreeAndNil(FFontStack);
  FreeAndNil(FCurrFont);
  FreeAndNil(FAttrList);
  inherited Destroy;
end;

{ Stores a font in the workbook's font list. Does not allow duplicates. }
function TsHTMLAnalyzer.AddFont(AFont: TsFont): Integer;
var
  fnt: TsFont;
begin
  // Is the font already stored in the workbook's font list?
  Result := FWorkbook.FindFont(AFont.FontName, AFont.Size, AFont.Style, AFont.Color, AFont.Position);
  if Result = -1 then
  begin
    // No. Create a new font, add it to the list, and return the new index.
    fnt := TsFont.Create;
    fnt.CopyOf(AFont);
    Result := FWorkbook.AddFont(fnt);
  end;
end;

procedure TsHTMLAnalyzer.AddRichTextParam(AFont: TsFont;
  AHyperlinkIndex: Integer = -1);
var
  len: Integer;
  fntIndex: Integer;
  n: Integer;
begin
  n := Length(FRichTextParams);
  len := UTF8Length(FPlainText);
  fntIndex := AddFont(AFont);
  if (n > 0) and (FRichTextparams[n-1].FirstIndex = len+1) then
  begin
    // Avoid adding another rich-text parameter for the same text location:
    // Update the previous one
    FRichTextParams[n-1].FontIndex := fntIndex;
    FRichTextParams[n-1].HyperlinkIndex := AHyperlinkIndex;
  end else
  begin
    // Add a new rich-text parameter
    SetLength(FRichTextParams, n+1);
    FRichTextParams[n].FirstIndex := len + 1;
    FRichTextParams[n].FontIndex := fntIndex;
    FRichTextParams[n].HyperlinkIndex := AHyperlinkIndex;
  end;
end;

procedure TsHTMLAnalyzer.ProcessFontRestore;
var
  fntIndex: Integer;
begin
  fntIndex := FFontStack.Pop;
  if fntIndex > -1 then
  begin
    FCurrFont.CopyOf(FWorkbook.GetFont(fntIndex));
    AddRichTextParam(FCurrFont);
  end;
end;

procedure TsHTMLAnalyzer.ReadFont(AFont: TsFont);
const
  FACTOR = 1.2;
  MIN_FONTSIZE = 6;
var
  idx: Integer;
  L: TStringList;
  i, ip, im: Integer;
  s: String;
  f: Double;
  defFntSize: Single;
begin
  idx := FAttrList.IndexOfName('font-family');    // style tag
  if idx = -1 then
    idx := FAttrList.IndexOfName('face');         // html tag
  if idx > -1 then begin
    L := TStringList.Create;
    try
      L.StrictDelimiter := true;
      L.DelimitedText := FAttrList[idx].Value;
      AFont.FontName := L[0];
    finally
      L.Free;
    end;
  end;

  idx := FAttrList.IndexOfName('font-size');
  if idx = -1 then
    idx := FAttrList.IndexOfName('size');
  if idx > -1 then begin
    defFntSize := FWorkbook.GetDefaultFont.Size;
    s := FAttrList[idx].Value;
    case s of
      'medium',   '3' : AFont.Size := defFntSize;
      'large',    '4' : AFont.Size := defFntSize*FACTOR;
      'x-large',  '5' : AFont.Size := defFntSize*FACTOR*FACTOR;
      'xx-large', '6' : AFont.Size := defFntSize*FACTOR*FACTOR*FACTOR;
      'small',    '2' : AFont.Size := Max(MIN_FONTSIZE, defFntSize/FACTOR);
      'x-small'       : AFont.Size := Max(MIN_FONTSIZE, defFntSize/FACTOR/FACTOR);
      'xx-small', '1' : AFont.Size := Max(MIN_FONTSIZE, defFntSize/FACTOR/FACTOR/FACTOR);
      'larger'        : AFont.Size := AFont.Size * FACTOR;
      'smaller'       : AFont.Size := Max(MIN_FONTSIZE, AFont.Size / FACTOR);
      else
        if s[1] in ['+', '-'] then
        begin
          TryStrToInt(s, i);
          AFont.Size := defFntSize * IntPower(FACTOR, i);
        end else
        begin
          i := 0;
          im := 0;
          ip := pos('%', s);
          if ip = 0 then begin
            im := pos('rem', s);
            if im = 0 then
              im := pos('em', s);
          end;
          if (ip > 0) then i := ip else
            if (im > 0) then i := im;
          if i > 0 then
          begin
            s := copy(s, 1, i-1);
            if TryStrToFloat(s, f, FPointSeparatorSettings) then
            begin
              if ip > 0 then f := f * 0.01;
              AFont.Size := Max(MIN_FONTSIZE, abs(f) * defFntSize);
            end;
          end else
            AFont.Size := Max(MIN_FONTSIZE, HTMLLengthStrToPts(s));
        end;
    end;
  end;

  idx := FAttrList.IndexOfName('font-style');
  if idx > -1 then
    case FAttrList[idx].Value of
      'normal'  : Exclude(AFont.Style, fssItalic);
      'italic'  : Include(AFont.Style, fssItalic);
      'oblique' : Include(AFont.Style, fssItalic);
    end;

  idx := FAttrList.IndexOfName('font-weight');
  if idx > -1 then
  begin
    s := FAttrList[idx].Value;
    if TryStrToInt(s, i) and (i >= 700) then Include(AFont.Style, fssBold);
  end;

  idx := FAttrList.IndexOfName('text-decoration');
  if idx > -1 then
  begin
    s := FAttrList[idx].Value;
    if pos('underline', s) <> 0 then Include(AFont.Style, fssUnderline);
    if pos('line-through', s) <> 0 then Include(AFont.Style, fssStrikeout);
  end;

  idx := FAttrList.IndexOfName('color');
  if idx > -1 then
    AFont.Color := HTMLColorStrToColor(FAttrList[idx].Value);
end;

procedure TsHTMLAnalyzer.TagFoundHandler(NoCaseTag, ActualTag: String);
begin
  case NoCaseTag[2] of
    'B': case NoCaseTag of
           '<B>'  : begin
                      FFontStack.Push(AddFont(FCurrFont));
                      Include(FCurrFont.Style, fssBold);
                      AddRichTextParam(FCurrFont);
                    end;
           '<BR>',
           '<BR/>': FPlainText := FPlainText + FPS_LINE_ENDING;
           else     if (pos('<BR ', NoCaseTag) = 1) then
                      FPlainText := FPlainText + FPS_LINE_ENDING;
         end;
    'D': if (NoCaseTag = '<DEL>') then
         begin
           FFontStack.Push(AddFont(FCurrFont));
           Include(FCurrFont.Style, fssStrikeout);
           AddRichTextParam(FCurrFont);
         end;
    'E': if (NoCaseTag = '<EM>') then
         begin
           FFontStack.Push(AddFont(FCurrFont));
           Include(FCurrFont.Style, fssItalic);
           AddRichTextParam(FCurrFont);
         end;
    'F': if (pos('<FONT ', NoCaseTag) = 1) then
         begin
           FFontStack.Push(AddFont(FCurrFont));
           FAttrList.Parse(ActualTag);
           ReadFont(FCurrFont);
           AddRichTextparam(FCurrFont);
         end;
    'I': case NoCaseTag of
           '<I>'  : begin
                      FFontStack.Push(AddFont(FCurrFont));
                      Include(FCurrFont.Style, fssItalic);
                      AddRichTextParam(FCurrFont);
                    end;
           '<INS>': begin
                      FFontStack.Push(AddFont(FCurrFont));
                      Include(FCurrFont.Style, fssUnderline);
                      AddRichTextParam(FCurrFont);
                    end;
         end;
    'S': case NoCaseTag of
           '<S>'  : begin
                      FFontStack.Push(AddFont(FCurrFont));
                      Include(FCurrFont.Style, fssStrikeout);
                      AddRichTextParam(FCurrFont);
                    end;
           '<STRONG>':begin
                       FFontStack.Push(AddFont(FCurrFont));
                       Include(FCurrFont.Style, fssBold);
                       AddRichTextParam(FCurrFont);
                    end;
           '<SUB>': begin
                      FFontStack.Push(AddFont(FCurrFont));
                      FCurrFont.Position := fpSubscript;
                      AddRichTextParam(FCurrFont);
                    end;
           '<SUP>': begin
                      FFontStack.Push(AddFont(FCurrFont));
                      FCurrFont.Position := fpSuperscript;
                      AddRichTextParam(FCurrFont);
                    end;
         end;
    'U': if (NoCaseTag = '<U>') then
         begin
           FFontStack.Push(AddFont(FCurrFont));
           Include(FCurrFont.Style, fssUnderline);
           AddRichTextParam(FCurrFont);
         end;
    '/': case NoCaseTag[3] of
           'B': if (NoCaseTag) = '</B>' then ProcessFontRestore;
           'D': if (NoCaseTag) = '</DEL>' then ProcessFontRestore;
           'E': if (NoCaseTag) = '</EM>' then ProcessFontRestore;
           'F': if (NoCaseTag) = '</FONT>' then ProcessFontRestore;
           'I': if (NoCaseTag = '</I>') or (NoCaseTag = '</INS>') then ProcessFontRestore;
           'S': if (NoCaseTag = '</S>') or (NoCaseTag = '</STRONG>') or
                   (NoCaseTag = '</SUB>') or (NoCaseTag = '</SUP>') then ProcessFontRestore;
           'U': if (NoCaseTag = '</U>') then ProcessFontRestore;
         end;
  end;
end;

procedure TsHTMLAnalyzer.TextFoundHandler(AText: String);
begin
  if not FPreserveSpaces then
    AText := CleanHTMLString(AText) else
    AText := RemoveHTMLEntities(AText);
  if AText <> '' then
  begin
    if FPlainText = '' then
      FPlainText := AText
    else
      FPlainText := FPlainText + AText;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Extracts rich-text parameters out of an html-formatted string and returns the
  plain text

  @@param AWorkbook          Workbook in which the cell with this text will reside
  @@param AFont              Standard font used in the cell in which this text
                             will reside
  @@param AHTMLText          Input Text with html tags to be applied to the cell
  @@param APlainText         AHtmlText after removal of HTML tags, pure text
  @@param ARichtTextParams   Rich-text parameters corresponding to the embedded
                             html tags
-------------------------------------------------------------------------------}
procedure HTMLToRichText(AWorkbook: TsWorkbook; AFont: TsFont;
  const AHTMLText: String; out APlainText: String;
  out ARichTextParams: TsRichTextParams);
var
  analyzer: TsHTMLAnalyzer;
  j: Integer;
  len: Integer;
  nrtp: Integer;
begin
  analyzer := TsHTMLAnalyzer.Create(AWorkbook, AFont, AHTMLText + '<end>');
  try
    analyzer.PreserveSpaces := true;
    analyzer.Exec;
    APlainText := analyzer.PlainText;
    nrtp := Length(analyzer.RichTextParams);
    if nrtp > 0 then begin
      // Avoid rich-text parameters beyond text end.
      len := UTF8Length(APlainText);
      if analyzer.RichTextParams[nrtp-1].FirstIndex > len then
        dec(nrtp);
      // Copy rich-text parameters to output
      SetLength(ARichTextParams, nrtp);
      for j:=0 to nrtp-1 do
        ARichTextParams[j] := analyzer.RichTextParams[j];
    end else
      SetLength(ARichTextParams, 0);
  finally
    analyzer.Free;
  end;
end;


{==============================================================================}
{                         Rich-text-to-HTML conversion                         }
{==============================================================================}

type
  TsHTMLComposer = class
  private
    FPointSeparatorSettings: TFormatSettings;
    FWorkbook: TsWorkbook;
    FBaseFont: TsFont;
    FPlainText: String;
    FRichTextParams: TsRichTextParams;
    FPrefix: String;
    FTagCase: TsTagCase;
    procedure GetFontsFromWorkbook(out AFonts: TsFontArray);
    function GetTextOfRichTextParam(AIndex: Integer): String;
  protected
    function FixTagCase(ATag: String): String;
  public
    constructor Create(AWorkbook: TsWorkbook; AFont: TsFont; APrefix: String = '';
      ATagCase: TsTagCase = tcLowercase);
    function Exec(const APlainText: String; const ARichTextParams: TsRichTextParams): String;
  end;

constructor TsHTMLComposer.Create(AWorkbook: TsWorkbook; AFont: TsFont;
  APrefix: String = ''; ATagCase: TsTagCase = tcLowercase);
begin
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';
  FWorkbook := AWorkbook;
  FBaseFont := AFont;
  FPrefix := APrefix;
  FTagCase := ATagCase;
end;

function TsHTMLComposer.Exec(const APlainText: String;
  const ARichTextParams: TsRichTextParams): String;
type
  TChangeFlag = (cfFontName, cfFontSize, cfFontColor);
const
  EPS = 1E-3;
var
  i: Integer;
  prevFnt, currFnt: TsFont;
  chgFlags: set of TChangeFlag;
  openingTag, closingTag: String;
  fonts: TsFontArray;
  tag: String;
begin
  if Length(ARichTextParams) = 0 then
  begin
    Result := FPlainText;
    exit;
  end;

  FRichTextParams := ARichTextParams;
  FPlainText := APlainText;

  prevFnt := TsFont.Create;
  prevFnt.CopyOf(FBaseFont);

  if FRichTextParams[0].FirstIndex > 1 then
    Result := GetTextOfRichTextParam(-1) else
    Result := '';

  GetFontsFromWorkbook(fonts);
  for i:=0 to High(FRichTextParams) do
  begin
    currFnt := fonts[i];
    openingTag := '';
    closingTag := '';
    if not SameFont(currFnt, prevFnt) then
    begin
      chgFlags := [];
      if not SameText(prevFnt.FontName, currFnt.FontName) then
        Include(chgFlags, cfFontName);
      if not SameValue(currFnt.Size, prevFnt.Size, EPS) then
        Include(chgFlags, cfFontSize);
      if currFnt.Color <> prevFnt.Color then
        Include(chgFlags, cfFontColor);

      if [cfFontName, cfFontSize, cfFontColor] * chgFlags <> [] then
      begin
        tag := FixTagCase('font');
        openingTag := '<' + tag;
        if cfFontName in chgFlags then
        begin
          openingTag := openingTag + ' ' + FPrefix + FixTagCase('face') +
            '="' + UnquoteStr(currFnt.FontName) + '"';
          prevFnt.FontName := currFnt.FontName;
        end;
        if cfFontSize in chgFlags then
        begin
          openingTag := openingTag + ' ' + FPrefix + FixTagCase('size') +
            '="' + Format('%.gpt', [currFnt.Size], FPointSeparatorSettings) + '"';
          prevFnt.Size := currFnt.Size;
        end;
        if cfFontColor in chgFlags then
        begin
          openingTag := openingTag + ' ' + FPrefix + FixTagCase('color') +
            '="' + ColorToHTMLColorStr(currFnt.Color) + '"';
          prevFnt.Size := currFnt.Color;
        end;
        openingTag := openingTag + '>';
        closingTag :='</' + tag + '>' + closingTag;
      end;

      if (fssBold in currFnt.Style) then
      begin
        tag := FixTagCase('b');
        openingTag := openingTag + '<' + tag + '>';
        closingTag := '</' + tag + '>' + closingTag;
        prevFnt.Style := prevFnt.Style + [fssBold];
      end else
        prevFnt.Style := prevFnt.Style - [fssBold];;

      if (fssItalic in currFnt.Style) then
      begin
        tag := FixTagCase('i');
        openingTag := openingTag + '<' + tag + '>';
        closingTag := '</' + tag + '>' + closingTag;
        prevFnt.Style := prevFnt.Style + [fssItalic];
      end else
        prevFnt.Style := prevFnt.Style - [fssItalic];

      if (fssUnderline in currFnt.Style) then
      begin
        tag := FixTagCase('u');
        openingTag := openingTag + '<' + tag + '>';
        closingTag := '</' + tag + '>' + closingTag;
        prevFnt.Style := prevFnt.Style + [fssUnderline];
      end else
        prevFnt.Style := prevFnt.Style - [fssUnderline];

      if (fssStrikeout in currFnt.Style) then
      begin
        tag := FixTagCase('s');
        openingTag := openingTag + '<' + tag + '>';
        closingTag := '</' + tag + '>' + closingTag;
        prevFnt.Style := prevFnt.Style + [fssStrikeout];
      end else
        prevFnt.Style := prevFnt.Style - [fssStrikeout];

      if currFnt.Position <> prevFnt.Position then
      begin
        if currFnt.Position = fpSuperscript then
        begin
          tag := FixTagCase('sup');
          openingTag := openingTag + '<' + tag + '>';
          closingTag := '</' + tag + '>' + closingTag;
          currFnt.Position := fpSuperscript;
        end else
        if currFnt.Position = fpSubscript then
        begin
          tag := FixTagCase('sub');
          openingTag := openingTag + '<' + tag + '>';
          closingTag := '</' + tag + '>' + closingTag;
          currFnt.Position := fpSubscript;
        end else
          currFnt.Position := fpNormal;
      end;
    end;

    // Add the node text with opening and closing tags (reverse order as opening!)
    Result := Result + openingTag + GetTextOfRichTextParam(i) + closingTag;
  end;  // for
end;

function TsHTMLComposer.FixTagCase(ATag: String): String;
begin
  case FTagCase of
    tcLowercase:
      Result := Lowercase(ATag);
    tcUppercase:
      Result := Uppercase(ATag);
    tcProperCase:
      begin
        Result := Lowercase(ATag);
        Result[1] := UpCase(Result[1]);
      end;
  end;
end;

procedure TsHTMLComposer.GetFontsFromWorkbook(out AFonts: TsFontArray);
var
  i: Integer;
begin
  SetLength(AFonts, Length(FRichTextParams));
  for i:=0 to High(AFonts) do
    AFonts[i] := FWorkbook.GetFont(FRichTextParams[i].FontIndex);
end;

function TsHTMLComposer.GetTextOfRichTextParam(AIndex: Integer): String;
var
  p1, p2: Integer;
begin
  if AIndex = -1 then
    Result := UTF8Copy(FPlainText, 1, FRichTextParams[0].FirstIndex-1)
  else
  if AIndex <= High(FRichTextParams) then
  begin
    p1 := FRichTextParams[AIndex].FirstIndex;
    if AIndex < High(FRichTextparams) then
      p2 := FRichTextParams[AIndex+1].FirstIndex else
      p2 := UTF8Length(FPlainText) + 1;
    Result := UTF8Copy(FPlaiNText, p1, p2-p1);
  end else
    Result := '';
end;


{@@ ----------------------------------------------------------------------------
  Constructs a html-coded string from a plain text string and
  rich-text parameters
-------------------------------------------------------------------------------}
procedure RichTextToHTML(AWorkbook: TsWorkbook; AFont: TsFont;
  const APlainText: String; const ARichTextParams: TsRichTextParams;
  out AHTMLText: String; APrefix: String = ''; ATagCase: TsTagCase = tcLowercase);
var
  composer: TsHTMLComposer;
begin
  if Length(ARichTextParams) = 0 then
    AHTMLText := APlainText
  else
  begin
    composer := TsHTMLComposer.Create(AWorkbook, AFont, APrefix, ATagCase);
    try
      AHTMLText := composer.Exec(APlainText, ARichTextParams);
    finally
      composer.Free;
    end;
  end;
end;

end.

