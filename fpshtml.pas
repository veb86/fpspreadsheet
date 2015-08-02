unit fpsHTML;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fasthtmlparser,
  fpstypes, fpspreadsheet, fpsReaderWriter, fpsHTMLUtils;

type
  TsHTMLTokenKind = (htkTABLE, htkTR, htkTH, htkTD, htkDIV, htkSPAN, htkP);
 {
  TsHTMLToken = class
    Kind: TsHTMLTokenKind;
    Parent: TsHTMLToken;
    Children
}
  TsHTMLReader = class(TsCustomSpreadReader)
  private
    FFormatSettings: TFormatSettings;
    parser: THTMLParser;
    FInTable: Boolean;
    FInSubTable: Boolean;
    FInCell: Boolean;
    FInSpan: Boolean;
    FInA: Boolean;
    FInHeader: Boolean;
    FTableCounter: Integer;
    FCurrRow, FCurrCol: LongInt;
    FCelLText: String;
    FAttrList: TsHTMLAttrList;
    FColSpan, FRowSpan: Integer;
    procedure ExtractMergedRange;
    procedure TagFoundHandler(NoCaseTag, ActualTag: string);
    procedure TextFoundHandler(AText: String);
  protected
    procedure AddCell(ARow, ACol: LongInt; AText: String);
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    procedure ReadFromStream(AStream: TStream); override;
    procedure ReadFromStrings(AStrings: TStrings); override;
  end;

  TsHTMLWriter = class(TsCustomSpreadWriter)
  private
    FPointSeparatorSettings: TFormatSettings;
//    function CellFormatAsString(ACell: PCell; ForThisTag: String): String;
    function CellFormatAsString(AFormat: PsCellFormat; ATagName: String): String;
    function GetBackgroundAsStyle(AFill: TsFillPattern): String;
    function GetBorderAsStyle(ABorder: TsCellBorders; const ABorderStyles: TsCellBorderStyles): String;
    function GetColWidthAsAttr(AColIndex: Integer): String;
    function GetDefaultHorAlignAsStyle(ACell: PCell): String;
    function GetFontAsStyle(AFontIndex: Integer): String;
    function GetGridBorderAsStyle: String;
    function GetHorAlignAsStyle(AHorAlign: TsHorAlignment): String;
    function GetMergedRangeAsStyle(AMergeBase: PCell): String;
    function GetRowHeightAsAttr(ARowIndex: Integer): String;
    function GetTextRotationAsStyle(ATextRot: TsTextRotation): String;
    function GetVertAlignAsStyle(AVertAlign: TsVertAlignment): String;
    function GetWordWrapAsStyle(AWordWrap: Boolean): String;
    function IsHyperlinkTarget(ACell: PCell; out ABookmark: String): Boolean;
    procedure WriteBody(AStream: TStream);
    procedure WriteStyles(AStream: TStream);
    procedure WriteWorksheet(AStream: TStream; ASheet: TsWorksheet);

  protected
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); override;
    procedure WriteError(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TsErrorValue; ACell: PCell); override;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    procedure WriteToStream(AStream: TStream); override;
    procedure WriteToStrings(AStrings: TStrings); override;
  end;

  TsHTMLParams = record
    TableIndex: Integer;             // R: Index of the table in the HTML file
    SheetIndex: Integer;             // W: Index of the sheet to be written
    ShowRowColHeaders: Boolean;      // RW: Show row/column headers
    DetectContentType: Boolean;      // R: try to convert strings to content types
    NumberFormat: String;            // W: if empty write numbers like in sheet, otherwise use this format
    AutoDetectNumberFormat: Boolean; // R: automatically detects decimal/thousand separator used in numbers
    TrueText: String;                // RW: String for boolean TRUE
    FalseText: String;               // RW: String for boolean FALSE
    FormatSettings: TFormatSettings; // RW: add'l parameters for conversion
  end;

var
  HTMLParams: TsHTMLParams = (
    TableIndex: -1;                  // -1 = all tables
    SheetIndex: -1;                  // -1 = active sheet, MaxInt = all sheets
    ShowRowColHeaders: false;
    DetectContentType: true;
    NumberFormat: '';
    AutoDetectNumberFormat: true;
    TrueText: 'TRUE';
    FalseText: 'FALSE';
  {%H-});

implementation

uses
  LazUTF8, URIParser, StrUtils,
  fpsUtils, fpsNumFormat;
                   (*
type
  THTMLEntity = record
    E: String;
    Ch: String;
  end;

const
  HTMLEntities: array[0..251] of THTMLEntity = (
  // A
    (E: 'Acirc';  Ch: 'Â'),  // 0
    (E: 'acirc';  Ch: 'â'),
    (E: 'acute';  Ch: '´'),
    (E: 'AElig';  Ch: 'Æ'),
    (E: 'aelig';  Ch: 'æ'),
    (E: 'Agrave'; Ch: 'À'),
    (E: 'agrave'; Ch: 'à'),
    (E: 'alefsym';Ch: 'ℵ'),
    (E: 'Alpha';  Ch: 'Α'),
    (E: 'alpha';  Ch: 'α'),
    (E: 'amp';    Ch: '&'),  // 10
    (E: 'and';    Ch: '∧'),
    (E: 'ang';    Ch: '∠'),
    (E: 'apos';   Ch: ''''),
    (E: 'Aring';  Ch: 'Å'),
    (E: 'aring';  Ch: 'å'),
    (E: 'asymp';  Ch: '≈'),
    (E: 'Atilde'; Ch: 'Ã'),
    (E: 'atilde'; Ch: 'ã'),
    (E: 'Auml';   Ch: 'Ä'),
    (E: 'auml';   Ch: 'ä'),  // 20
  // B
    (E: 'bdquo';  Ch: '„'),  // 21
    (E: 'Beta';   Ch: 'Β'),
    (E: 'beta';   Ch: 'β'),
    (E: 'brvbar'; Ch: '¦'),
    (E: 'bull';   Ch: '•'),
  // C
    (E: 'cap';    Ch: '∩'),  // 26
    (E: 'Ccedil'; Ch: 'Ç'),
    (E: 'ccedil'; Ch: 'ç'),
    (E: 'cedil';  Ch: '¸'),
    (E: 'cent';   Ch: '¢'),  // 39
    (E: 'Chi';    Ch: 'Χ'),
    (E: 'chi';    Ch: 'χ'),
    (E: 'circ';   Ch: 'ˆ'),
    (E: 'clubs';  Ch: '♣'),
    (E: 'cong';   Ch: '≅'),  // approximately equal
    (E: 'copy';   Ch: '©'),
    (E: 'crarr';  Ch: '↵'),  // carriage return
    (E: 'cup';    Ch: '∪'),
    (E: 'curren'; Ch: '¤'),
  // D
    (E: 'Dagger'; Ch: '‡'),  // 40
    (E: 'dagger'; Ch: '†'),
    (E: 'dArr';   Ch: '⇓'),  // wide down-arrow
    (E: 'darr';   Ch: '↓'),  // narrow down-arrow
    (E: 'deg';    Ch: '°'),
    (E: 'Delta';  Ch: 'Δ'),
    (E: 'delta';  Ch: 'δ'),
    (E: 'diams';  Ch: '♦'),
    (E: 'divide'; Ch: '÷'),
  // E
    (E: 'Eacute'; Ch: 'É'),
    (E: 'eacute'; Ch: 'é'),
    (E: 'Ecirc';  Ch: 'Ê'),
    (E: 'ecirc';  Ch: 'ê'),
    (E: 'Egrave'; Ch: 'È'),
    (E: 'egrave'; Ch: 'è'),
    (E: 'empty';  Ch: '∅'),
    (E: 'emsp';   Ch: ' '),  // Space character width of "m"
    (E: 'ensp';   Ch: ' '),  // Space character width of "n"
    (E: 'Epsilon';Ch: 'Ε'),  // capital epsilon
    (E: 'epsilon';Ch: 'ε'),
    (E: 'equiv';  Ch: '≡'),
    (E: 'Eta';    Ch: 'Η'),
    (E: 'eta';    Ch: 'η'),
    (E: 'ETH';    Ch: 'Ð'),
    (E: 'eth';    Ch: 'ð'),
    (E: 'Euml';   Ch: 'Ë'),
    (E: 'euml';   Ch: 'ë'),
    (E: 'euro';   Ch: '€'),
    (E: 'exist';  Ch: '∃'),
  // F
    (E: 'fnof';   Ch: 'ƒ'),
    (E: 'forall'; Ch: '∀'),
    (E: 'frac12'; Ch: '½'),
    (E: 'frac14'; Ch: '¼'),
    (E: 'frac34'; Ch: '¾'),
    (E: 'frasl';  Ch: '⁄'),
  // G
    (E: 'Gamma';  Ch: 'Γ'),
    (E: 'gamma';  Ch: 'γ'),
    (E: 'ge';     Ch: '≥'),
    (E: 'gt';     Ch: '>'),
  // H
    (E: 'hArr';   Ch: '⇔'),  // wide horizontal double arrow
    (E: 'harr';   Ch: '↔'),  // narrow horizontal double arrow
    (E: 'hearts'; Ch: '♥'),
    (E: 'hellip'; Ch: '…'),
  // I
    (E: 'Iacute'; Ch: 'Í'),
    (E: 'iacute'; Ch: 'í'),
    (E: 'Icirc';  Ch: 'Î'),
    (E: 'icirc';  Ch: 'î'),
    (E: 'iexcl';  Ch: '¡'),
    (E: 'Igrave'; Ch: 'Ì'),
    (E: 'igrave'; Ch: 'ì'),
    (E: 'image';  Ch: 'ℑ'),  //
    (E: 'infin';  Ch: '∞'),
    (E: 'int';    Ch: '∫'),
    (E: 'Iota';   Ch: 'Ι'),
    (E: 'iota';   Ch: 'ι'),
    (E: 'iquest'; Ch: '¿'),
    (E: 'isin';   Ch: '∈'),
    (E: 'Iuml';   Ch: 'Ï'),
    (E: 'iuml';   Ch: 'ï'),
  // K
    (E: 'Kappa';  Ch: 'Κ'),
    (E: 'kappa';  Ch: 'κ'),
  // L
    (E: 'Lambda'; Ch: 'Λ'),
    (E: 'lambda'; Ch: 'λ'),
    (E: 'lang';   Ch: '⟨'),  // Left-pointing angle bracket
    (E: 'laquo';  Ch: '«'),
    (E: 'lArr';   Ch: '⇐'),  // Left-pointing wide arrow
    (E: 'larr';   Ch: '←'),
    (E: 'lceil';  Ch: '⌈'),  // Left ceiling
    (E: 'ldquo';  Ch: '“'),
    (E: 'le';     Ch: '≤'),
    (E: 'lfloor'; Ch: '⌊'),  // Left floor
    (E: 'lowast'; Ch: '∗'),  // Low asterisk
    (E: 'loz';    Ch: '◊'),
    (E: 'lrm';    Ch: '‎'),  // Left-to-right mark
    (E: 'lsaquo'; Ch: '‹'),
    (E: 'lsquo';  Ch: '‘'),
    (E: 'lt';     Ch: '<'),
  // M
    (E: 'macr';   Ch: '¯'),
    (E: 'mdash';  Ch: '—'),
    (E: 'micro';  Ch: 'µ'),
    (E: 'middot'; Ch: '·'),
    (E: 'minus';  Ch: '−'),
    (E: 'Mu';     Ch: 'Μ'),
    (E: 'mu';     Ch: 'μ'),
  // N
    (E: 'nabla';  Ch: '∇'),
    (E: 'nbsp';   Ch: ' '),
    (E: 'ndash';  Ch: '–'),
    (E: 'ne';     Ch: '≠'),
    (E: 'ni';     Ch: '∋'),
    (E: 'not';    Ch: '¬'),
    (E: 'notin';  Ch: '∉'),  // math: "not in"
    (E: 'nsub';   Ch: '⊄'),  // math: "not a subset of"
    (E: 'Ntilde'; Ch: 'Ñ'),
    (E: 'ntilde'; Ch: 'ñ'),
    (E: 'Nu';     Ch: 'Ν'),
    (E: 'nu';     Ch: 'ν'),
  // O
    (E: 'Oacute'; Ch: 'Ó'),
    (E: 'oacute'; Ch: 'ó'),
    (E: 'Ocirc';  Ch: 'Ô'),
    (E: 'ocirc';  Ch: 'ô'),
    (E: 'OElig';  Ch: 'Œ'),
    (E: 'oelig';  Ch: 'œ'),
    (E: 'Ograve'; Ch: 'Ò'),
    (E: 'ograve'; Ch: 'ò'),
    (E: 'oline';  Ch: '‾'),
    (E: 'Omega';  Ch: 'Ω'),
    (E: 'omega';  Ch: 'ω'),
    (E: 'Omicron';Ch: 'Ο'),
    (E: 'omicron';Ch: 'ο'),
    (E: 'oplus';  Ch: '⊕'),  // Circled plus
    (E: 'or';     Ch: '∨'),
    (E: 'ordf';   Ch: 'ª'),
    (E: 'ordm';   Ch: 'º'),
    (E: 'Oslash'; Ch: 'Ø'),
    (E: 'oslash'; Ch: 'ø'),
    (E: 'Otilde'; Ch: 'Õ'),
    (E: 'otilde'; Ch: 'õ'),
    (E: 'otimes'; Ch: '⊗'),  // Circled times
    (E: 'Ouml';   Ch: 'Ö'),
    (E: 'ouml';   Ch: 'ö'),
  // P
    (E: 'para';   Ch: '¶'),
    (E: 'part';   Ch: '∂'),
    (E: 'permil'; Ch: '‰'),
    (E: 'perp';   Ch: '⊥'),
    (E: 'Phi';    Ch: 'Φ'),
    (E: 'phi';    Ch: 'φ'),
    (E: 'Pi';     Ch: 'Π'),
    (E: 'pi';     Ch: 'π'),  // lower-case pi
    (E: 'piv';    Ch: 'ϖ'),
    (E: 'plusmn'; Ch: '±'),
    (E: 'pound';  Ch: '£'),
    (E: 'Prime';  Ch: '″'),
    (E: 'prime';  Ch: '′'),
    (E: 'prod';   Ch: '∏'),
    (E: 'prop';   Ch: '∝'),
    (E: 'Psi';    Ch: 'Ψ'),
    (E: 'psi';    Ch: 'ψ'),
  // Q
    (E: 'quot';   Ch: '"'),
  // R
    (E: 'radic';  Ch: '√'),
    (E: 'rang';   Ch: '⟩'),  // right-pointing angle bracket
    (E: 'raquo';  Ch: '»'),
    (E: 'rArr';   Ch: '⇒'),
    (E: 'rarr';   Ch: '→'),
    (E: 'rceil';  Ch: '⌉'),  // right ceiling
    (E: 'rdquo';  Ch: '”'),
    (E: 'real';   Ch: 'ℜ'),  // R in factura
    (E: 'reg';    Ch: '®'),
    (E: 'rfloor'; Ch: '⌋'),  // Right floor
    (E: 'Rho';    Ch: 'Ρ'),
    (E: 'rho';    Ch: 'ρ'),
    (E: 'rlm';    Ch: ''),   // right-to-left mark
    (E: 'rsaquo'; Ch: '›'),
    (E: 'rsquo';  Ch: '’'),

  // S
    (E: 'sbquo';  Ch: '‚'),
    (E: 'Scaron'; Ch: 'Š'),
    (E: 'scaron'; Ch: 'š'),
    (E: 'sdot';   Ch: '⋅'),  // math: dot operator
    (E: 'sect';   Ch: '§'),
    (E: 'shy';    Ch: ''),   // conditional hyphen
    (E: 'Sigma';  Ch: 'Σ'),
    (E: 'sigma';  Ch: 'σ'),
    (E: 'sigmaf'; Ch: 'ς'),
    (E: 'sim';    Ch: '∼'),  // similar
    (E: 'spades'; Ch: '♠'),
    (E: 'sub';    Ch: '⊂'),
    (E: 'sube';   Ch: '⊆'),
    (E: 'sum';    Ch: '∑'),
    (E: 'sup';    Ch: '⊃'),
    (E: 'sup1';   Ch: '¹'),
    (E: 'sup2';   Ch: '²'),
    (E: 'sup3';   Ch: '³'),
    (E: 'supe';   Ch: '⊇'),
    (E: 'szlig';  Ch: 'ß'),
  //T
    (E: 'Tau';    Ch: 'Τ'),
    (E: 'tau';    Ch: 'τ'),
    (E: 'there4'; Ch: '∴'),
    (E: 'Theta';  Ch: 'Θ'),
    (E: 'theta';  Ch: 'θ'),
    (E: 'thetasym';Ch: 'ϑ'),
    (E: 'thinsp'; Ch: ' '),  // thin space
    (E: 'THORN';  Ch: 'Þ'),
    (E: 'thorn';  Ch: 'þ'),
    (E: 'tilde';  Ch: '˜'),
    (E: 'times';  Ch: '×'),
    (E: 'trade';  Ch: '™'),
  // U
    (E: 'Uacute'; Ch: 'Ú'),
    (E: 'uacute'; Ch: 'ú'),
    (E: 'uArr';   Ch: '⇑'),  // wide up-arrow
    (E: 'uarr';   Ch: '↑'),
    (E: 'Ucirc';  Ch: 'Û'),
    (E: 'ucirc';  Ch: 'û'),
    (E: 'Ugrave'; Ch: 'Ù'),
    (E: 'ugrave'; Ch: 'ù'),
    (E: 'uml';    Ch: '¨'),
    (E: 'upsih';  Ch: 'ϒ'),
    (E: 'Upsilon';Ch: 'Υ'),
    (E: 'upsilon';Ch: 'υ'),
    (E: 'Uuml';   Ch: 'Ü'),
    (E: 'uuml';   Ch: 'ü'),
  // W
    (E: 'weierp'; Ch: '℘'),  // Script Capital P; Weierstrass Elliptic Function
  // X
    (E: 'Xi';     Ch: 'Ξ'),
    (E: 'xi';     Ch: 'ξ'),
  // Y
    (E: 'Yacute'; Ch: 'Ý'),
    (E: 'yacute'; Ch: 'ý'),
    (E: 'yen';    Ch: '¥'),
    (E: 'Yuml';   Ch: 'Ÿ'),
    (E: 'yuml';   Ch: 'ÿ'),
  // Z
    (E: 'Zeta';   Ch: 'Ζ'),
    (E: 'zeta';   Ch: 'ζ'),
    (E: 'zwj';    Ch: ''),   // Zero-width joiner
    (E: 'zwnj';   Ch: ''),   // Zero-width non-joiner

    (E: '#160';   Ch: ' ')   // numerical value of "&nbsp;"
  );
                     *)
{==============================================================================}
{                             TsHTMLReader                                     }
{==============================================================================}

constructor TsHTMLReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FFormatSettings := HTMLParams.FormatSettings;
  ReplaceFormatSettings(FFormatSettings, FWorkbook.FormatSettings);
  FTableCounter := -1;
  FAttrList := TsHTMLAttrList.Create;
end;

destructor TsHTMLReader.Destroy;
begin
  FreeAndNil(FAttrList);
  FreeAndNil(parser);
  inherited Destroy;
end;

procedure TsHTMLReader.AddCell(ARow, ACol: LongInt; AText: String);
var
  cell: PCell;
  dblValue: Double;
  dtValue: TDateTime;
  boolValue: Boolean;
  nf: TsNumberFormat;
  decs: Integer;
  currSym: String;
  warning: String;
begin
  // Empty strings are blank cells -- nothing to do
  if (AText = '') then
    exit;

  cell := FWorksheet.AddCell(ARow, ACol);

  // Merged cells
  if (FColSpan > 0) or (FRowSpan > 0) then
    FWorksheet.MergeCells(ARow, ACol, ARow + FRowSpan, ACol + FColSpan);

  // Do not try to interpret the strings. --> everything is a LABEL cell.
  if not HTMLParams.DetectContentType then
  begin
    FWorksheet.WriteUTF8Text(cell, AText);
    exit;
  end;

  // Check for a NUMBER or CURRENCY cell
  if IsNumberValue(AText, HTMLParams.AutoDetectNumberFormat, FFormatSettings,
    dblValue, nf, decs, currSym, warning) then
  begin
    if currSym <> '' then
      FWorksheet.WriteCurrency(cell, dblValue, nfCurrency, decs, currSym)
    else
      FWorksheet.WriteNumber(cell, dblValue, nf, decs);
    if warning <> '' then
      FWorkbook.AddErrorMsg('Cell %s: %s', [GetCellString(ARow, ACol), warning]);
    exit;
  end;

  // Check for a DATE/TIME cell
  // No idea how to apply the date/time formatsettings here...
  if IsDateTimevalue(AText, FFormatSettings, dtValue, nf) then
  begin
    FWorksheet.WriteDateTime(cell, dtValue, nf);
    exit;
  end;

  // Check for a BOOLEAN cell
  if IsBoolValue(AText, HTMLParams.TrueText, HTMLParams.FalseText, boolValue) then
  begin
    FWorksheet.WriteBoolValue(cell, boolValue);
    exit;
  end;

  // What is left is handled as a TEXT cell
  FWorksheet.WriteUTF8Text(cell, AText);
end;

procedure TsHTMLReader.ExtractMergedRange;
var
  idx: Integer;
begin
  FColSpan := 0;
  FRowSpan := 0;
  idx := FAttrList.IndexOfName('colspan');
  if idx > -1 then
    FColSpan := StrToInt(FAttrList[idx].Value) - 1;
  idx := FAttrList.IndexOfName('rowspan');
  if idx > -1 then
    FRowSpan := StrToInt(FAttrList[idx].Value) - 1;
  // -1 to compensate for correct determination of the range end cell
end;

procedure TsHTMLReader.ReadFromStream(AStream: TStream);
var
  list: TStringList;
begin
  list := TStringList.Create;
  try
    list.LoadFromStream(AStream);
    ReadFromStrings(list);
    if FWorkbook.GetWorksheetCount = 0 then
    begin
      FWorkbook.AddErrorMsg('Requested table not found, or no tables in html file');
      FWorkbook.AddWorksheet('Dummy');
    end;
  finally
    list.Free;
  end;
end;

procedure TsHTMLReader.ReadFromStrings(AStrings: TStrings);
begin
  // Create html parser
  FreeAndNil(parser);
  parser := THTMLParser.Create(AStrings.Text);
  parser.OnFoundTag := @TagFoundHandler;
  parser.OnFoundText := @TextFoundHandler;
  // Execute the html parser
  parser.Exec;
end;

procedure TsHTMLReader.TagFoundHandler(NoCaseTag, ActualTag: string);
begin
  if pos('<TABLE', NoCaseTag) = 1 then
  begin
    inc(FTableCounter);
    if HTMLParams.TableIndex < 0 then  // all tables
    begin
      FWorksheet := FWorkbook.AddWorksheet(Format('Table #%d', [FTableCounter+1]));
      FInTable := true;
      FCurrRow := -1;
      FCurrCol := -1;
    end else
    if FTableCounter = HTMLParams.TableIndex then
    begin
      FWorksheet := FWorkbook.AddWorksheet(Format('Table #%d', [FTableCounter+1]));
      FInTable := true;
      FCurrRow := -1;
      FCurrCol := -1;
    end;
  end else
  if ((NoCaseTag = '<TR>') or (pos('<TR ', NoCaseTag) = 1)) and FInTable then
  begin
    inc(FCurrRow);
    FCurrCol := -1;
  end else
  if ((NoCaseTag = '<TD>') or (pos('<TD ', NoCaseTag) = 1)) and FInTable then
  begin
    FInCell := true;
    inc(FCurrCol);
    FCellText := '';
    FAttrList.Parse(ActualTag);
    ExtractMergedRange;
  end else
  if ((NoCaseTag = '<TH>') or (pos('<TH ', NoCaseTag) = 1)) and FInTable then
  begin
    FInCell := true;
    FCellText := '';
  end else
  if pos('<SPAN', NoCaseTag) = 1 then
  begin
    if FInCell then
      FInSpan := true;
  end else
  if pos('<A', NoCaseTag) = 1 then
  begin
    if FInCell then
      FInA := true
  end else
  if (pos('<H', NoCaseTag) = 1) and (NoCaseTag[3] in ['1', '2', '3', '4', '5', '6']) then
  begin
    if FInCell then
      FInHeader := true;
  end else
  if ((NoCaseTag = '<BR>') or (pos('<BR ', NoCaseTag) = 1)) and FInCell then
    FCellText := FCellText + LineEnding
  else
    case NoCaseTag of
      '</TABLE>':
        if FInTable then FInTable := false;
      '</TD>', '</TH>':
        if FInCell then
        begin
//          inc(FCurrCol);
          while FWorksheet.isMerged(FWorksheet.FindCell(FCurrRow, FCurrRow)) do
            inc(FCurrRow);
            {
          if FWorksheet.IsMerged(FWorksheet.FindCell(FCurrRow, FCurrCol)) then
          begin
            repeat
              inc(FCurrRow);
            until not FWorksheet.IsMerged(FWorksheet.FindCell(FCurrRow, FCurrCol));
            dec(FCurrCol);
          end;
          }
          AddCell(FCurrRow, FCurrCol, FCellText);
          FInCell := false;
        end;
      '</A>':
        if FInCell then FInA := false;
      '</SPAN>':
        if FInCell then FInSpan := false;
      '<H1/>', '<H2/>', '<H3/>', '<H4/>', '<H5/>', '<H6/>':
        if FinCell then FInHeader := false;
      '<TR/>', '<TR />':                     // empty rows
        if FInTable then inc(FCurrRow);
      '<TD/>', '<TD />', '<TH/>', '<TH />':  // empty cells
        if FInCell then
          inc(FCurrCol);
    end;
end;

procedure TsHTMLReader.TextFoundHandler(AText: String);
begin
  if FInCell then
  begin
    AText := CleanHTMLString(AText);
    if AText <> '' then
    begin
      if FCellText = '' then
        FCellText := AText
      else
        FCellText := FCellText + ' ' + AText;
    end;
  end;
end;

{==============================================================================}
{                             TsHTMLWriter                                     }
{==============================================================================}
constructor TsHTMLWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  // No design limiations in table size
  // http://stackoverflow.com/questions/4311283/max-columns-in-html-table
  FLimitations.MaxColCount := MaxInt;
  FLimitations.MaxRowCount := MaxInt;
end;

destructor TsHTMLWriter.Destroy;
begin
  inherited Destroy;
end;
               (*
function TsHTMLWriter.CellFormatAsString(ACell: PCell; ForThisTag: String): String;
var
  fmt: PsCellFormat;
begin
  Result := '';
  if ACell <> nil then
    fmt := FWorkbook.GetPointerToCellFormat(ACell^.FormatIndex)
  else
    fmt := nil;
  case ForThisTag of
    'td':
      if ACell = nil then
      begin
        Result := 'border-collapse:collapse;';
        if soShowGridLines in FWorksheet.Options then
          Result := Result + GetGridBorderAsStyle;
      end else
      begin
        if (uffBackground in fmt^.UsedFormattingFields) then
          Result := Result + GetBackgroundAsStyle(fmt^.Background);
        if (uffFont in fmt^.UsedFormattingFields) then
          Result := Result + GetFontAsStyle(fmt^.FontIndex);
        if (uffTextRotation in fmt^.UsedFormattingFields) then
          Result := Result + GetTextRotationAsStyle(fmt^.TextRotation);
        if (uffHorAlign in fmt^.UsedFormattingFields) and (fmt^.HorAlignment <> haDefault) then
          Result := Result + GetHorAlignAsStyle(fmt^.HorAlignment)
        else
          case ACell^.ContentType of
            cctNumber    : Result := Result + GetHorAlignAsStyle(haRight);
            cctDateTime  : Result := Result + GetHorAlignAsStyle(haLeft);
            cctBool      : Result := Result + GetHorAlignAsStyle(haCenter);
            else           Result := Result + GetHorAlignAsStyle(haLeft);
          end;
        if (uffVertAlign in fmt^.UsedFormattingFields) then
          Result := Result + GetVertAlignAsStyle(fmt^.VertAlignment);
        if (uffBorder in fmt^.UsedFormattingFields) then
          Result := Result + GetBorderAsStyle(fmt^.Border, fmt^.BorderStyles)
        else begin
          if soShowGridLines in FWorksheet.Options then
            Result := Result + GetGridBorderAsStyle;
        end;
        if (uffFont in fmt^.UsedFormattingFields) then
          Result := Result + GetFontAsStyle(fmt^.FontIndex);   {
        if (uffTextRotation in fmt^.UsedFormattingFields) then
          Result := Result + GetTextRotation(fmt^.TextRotation);}
        Result := Result + GetWordwrapAsStyle(uffWordwrap in fmt^.UsedFormattingFields);
      end;
    'div', 'p':
      begin
        if fmt = nil then
          exit;
        {
        if (uffHorAlign in fmt^.UsedFormattingFields) and (fmt^.HorAlignment <> haDefault) then
          Result := Result + GetHorAlignAsStyle(fmt^.HorAlignment)
        else
          case ACell^.ContentType of
            cctNumber    : Result := Result + GetHorAlignAsStyle(haRight);
            cctDateTime  : Result := Result + GetHorAlignAsStyle(haLeft);
            cctBool      : Result := Result + GetHorAlignAsStyle(haCenter);
            else           Result := Result + GetHorAlignAsStyle(haLeft);
          end;
        if (uffFont in fmt^.UsedFormattingFields) then
          Result := Result + GetFontAsStyle(fmt^.FontIndex);   {
        if (uffTextRotation in fmt^.UsedFormattingFields) then
          Result := Result + GetTextRotation(fmt^.TextRotation);}
        Result := Result + GetWordwrapAsStyle(uffWordwrap in fmt^.UsedFormattingFields);
        }
      end;
  end;
  if Result <> '' then
    Result := ' style="' + Result +'"';
end;
         *)
function TsHTMLWriter.CellFormatAsString(AFormat: PsCellFormat; ATagName: String): String;
begin
  Result := '';

  if (uffBackground in AFormat^.UsedFormattingFields) then
    Result := Result + GetBackgroundAsStyle(AFormat^.Background);

  if (uffFont in AFormat^.UsedFormattingFields) then
    Result := Result + GetFontAsStyle(AFormat^.FontIndex);

  if (uffTextRotation in AFormat^.UsedFormattingFields) then
    Result := Result + GetTextRotationAsStyle(AFormat^.TextRotation);

  if (uffHorAlign in AFormat^.UsedFormattingFields) and (AFormat^.HorAlignment <> haDefault) then
    Result := Result + GetHorAlignAsStyle(AFormat^.HorAlignment);

  if (uffVertAlign in AFormat^.UsedFormattingFields) then
    Result := Result + GetVertAlignAsStyle(AFormat^.VertAlignment);

  if (uffBorder in AFormat^.UsedFormattingFields) then
    Result := Result + GetBorderAsStyle(AFormat^.Border, AFormat^.BorderStyles);
  {
  else begin
    if soShowGridLines in FWorksheet.Options then
      Result := Result + GetGridBorderAsStyle;
  end;
  }

  Result := Result + GetWordwrapAsStyle(uffWordwrap in AFormat^.UsedFormattingFields);
end;

function TsHTMLWriter.GetBackgroundAsStyle(AFill: TsFillPattern): String;
begin
  Result := '';
  if AFill.Style = fsSolidFill then
    Result := 'background-color:' + ColorToHTMLColorStr(AFill.FgColor) + ';';
  // other fills not supported
end;

function TsHTMLWriter.GetBorderAsStyle(ABorder: TsCellBorders;
  const ABorderStyles: TsCellBorderStyles): String;
const
  BORDER_NAMES: array[TsCellBorder] of string = (
    'border-top',    // cbNorth
    'border-left',   // cbWest
    'border-right',  // cbEast
    'border-bottom', // cbSouth
    '',              // cbDiagUp
    ''               // cbDiagDown
  );
  LINESTYLE_NAMES: array[TsLineStyle] of string = (
    'thin solid',    // lsThin
    'medium solid',  // lsMedium
    'thin dashed',   // lsDashed
    'thin dotted',   // lsDotted
    'thick solid',   // lsThick,
    'double',        // lsDouble,
    '1px solid',     // lsHair
    'medium dashed', // lsMediumDash     --- not all available in HTML...
    'thin dashed',   // lsDashDot
    'medium dashed', // lsMediumDashDot
    'thin dotted',   // lsDashDotDot
    'medium dashed', // lsMediumDashDotDot
    'medium dashed'  // lsSlantedDashDot
  );
var
  cb: TsCellBorder;
  allEqual: Boolean;
  bs: TsCellBorderStyle;
begin
  Result := 'border-collape:collapse;';
  if ABorder = [cbNorth, cbEast, cbWest, cbSouth] then
  begin
    allEqual := true;
    bs := ABorderStyles[cbNorth];
    for cb in TsCellBorder do
    begin
      if bs.LineStyle <> ABorderStyles[cb].LineStyle then
      begin
        allEqual := false;
        break;
      end;
      if bs.Color <> ABorderStyles[cb].Color then
      begin
        allEqual := false;
        break;
      end;
    end;
    if allEqual then
    begin
      Result := 'border:' +
        LINESTYLE_NAMES[bs.LineStyle] + ' ' +
        ColorToHTMLColorStr(bs.Color) + ';';
      exit;
    end;
  end;

  for cb in TsCellBorder do
  begin
    if BORDER_NAMES[cb] = '' then
      continue;
    if cb in ABorder then
      Result := Result + BORDER_NAMES[cb] + ':' +
        LINESTYLE_NAMES[ABorderStyles[cb].LineStyle] + ' ' +
        ColorToHTMLColorStr(ABorderStyles[cb].Color) + ';';
  end;
end;

function TsHTMLWriter.GetColWidthAsAttr(AColIndex: Integer): String;
var
  col: PCol;
  w: Single;
  rLast: Cardinal;
begin
  if AColIndex < 0 then  // Row header column
  begin
    rLast := FWorksheet.GetLastRowIndex;
    w := Length(IntToStr(rLast)) + 2;
  end else
  begin
    w := FWorksheet.DefaultColWidth;
    col := FWorksheet.FindCol(AColIndex);
    if col <> nil then
      w := col^.Width;
  end;
  w := w * FWorkbook.GetDefaultFont.Size;
  Result:= Format(' width="%.1fpt"', [w], FPointSeparatorSettings);
end;

function TsHTMLWriter.GetDefaultHorAlignAsStyle(ACell: PCell): String;
begin
  Result := '';
  if ACell = nil then
    exit;
  case ACell^.ContentType of
    cctNumber  : Result := GetHorAlignAsStyle(haRight);
    cctDateTime: Result := GetHorAlignAsStyle(haRight);
    cctBool    : Result := GetHorAlignAsStyle(haCenter);
  end;
end;

function TsHTMLWriter.GetFontAsStyle(AFontIndex: Integer): String;
var
  font: TsFont;
begin
  font := FWorkbook.GetFont(AFontIndex);
  Result := Format('font-family:''%s'';font-size:%.1fpt;color:%s;', [
    font.FontName, font.Size, ColorToHTMLColorStr(font.Color)], FPointSeparatorSettings);
  if fssBold in font.Style then
    Result := Result + 'font-weight:700;';
  if fssItalic in font.Style then
    Result := Result + 'font-style:italic;';
  if [fssUnderline, fssStrikeout] * font.Style = [fssUnderline, fssStrikeout] then
    Result := Result + 'text-decoration:underline,line-through;'
  else
  if [fssUnderline, fssStrikeout] * font.Style = [fssUnderline] then
    Result := Result + 'text-decoration:underline;'
  else
  if [fssUnderline, fssStrikeout] * font.Style = [fssStrikeout] then
    Result := Result + 'text-decoration:line-through;';
end;

function TsHTMLWriter.GetGridBorderAsStyle: String;
begin
  if (soShowGridLines in FWorksheet.Options) then
    Result := 'border:1px solid lightgrey;'
  else
    Result := '';
end;

function TsHTMLWriter.GetHorAlignAsStyle(AHorAlign: TsHorAlignment): String;
begin
  case AHorAlign of
    haLeft   : Result := 'text-align:left;';
    haCenter : Result := 'text-align:center;';
    haRight  : Result := 'text-align:right;';
  end;
end;

function TsHTMLWriter.GetMergedRangeAsStyle(AMergeBase: PCell): String;
var
  r1, r2, c1, c2: Cardinal;
begin
  Result := '';
  FWorksheet.FindMergedRange(AMergeBase, r1, c1, r2, c2);
  if c1 <> c2 then
    Result := Result + ' colspan="' + IntToStr(c2-c1+1) + '"';
  if r1 <> r2 then
    Result := Result + ' rowspan="' + IntToStr(r2-r1+1) + '"';
end;

function TsHTMLWriter.GetRowHeightAsAttr(ARowIndex: Integer): String;
var
  h: Single;
  row: PRow;
begin
  h := FWorksheet.DefaultRowHeight;
  row := FWorksheet.FindRow(ARowIndex);
  if row <> nil then
    h := row^.Height;
  h := (h + ROW_HEIGHT_CORRECTION) * FWorkbook.GetDefaultFont.Size;
  Result := Format(' height="%.1fpt"', [h], FPointSeparatorSettings);
end;


function TsHTMLWriter.GetTextRotationAsStyle(ATextRot: TsTextRotation): String;
begin
  Result := '';
  case ATextRot of
    trHorizontal: ;
    rt90DegreeClockwiseRotation:
      Result := 'writing-mode:vertical-rl;transform:rotate(90deg);'; //-moz-transform: rotate(90deg);';
//      Result := 'writing-mode:vertical-rl;text-orientation:sideways-right;-moz-transform: rotate(-90deg);';
    rt90DegreeCounterClockwiseRotation:
      Result := 'writing-mode:vertical-rt;transform:rotate(-90deg);'; //-moz-transform: rotate(-90deg);';
//    Result := 'writing-mode:vertical-rt;text-orientation:sideways-left;-moz-transform: rotate(-90deg);';
    rtStacked:
      Result := 'writing-mode:vertical-rt;text-orientation:upright;';
  end;
end;

function TsHTMLWriter.GetVertAlignAsStyle(AVertAlign: TsVertAlignment): String;
begin
  case AVertAlign of
    vaTop    : Result := 'vertical-align:top;';
    vaCenter : Result := 'vertical-align:middle;';
    vaBottom : Result := 'vertical-align:bottom;';
  end;
end;

function TsHTMLWriter.GetWordwrapAsStyle(AWordwrap: Boolean): String;
begin
  if AWordwrap then
    Result := 'word-wrap:break-word;'
  else
    Result := 'white-space:nowrap;';
end;

function TsHTMLWriter.IsHyperlinkTarget(ACell: PCell; out ABookmark: String): Boolean;
var
  sheet: TsWorksheet;
  hyperlink: PsHyperlink;
  target, sh: String;
  i, r, c: Cardinal;
begin
  Result := false;
  if ACell = nil then
    exit;

  for i:=0 to FWorkbook.GetWorksheetCount-1 do
  begin
    sheet := FWorkbook.GetWorksheetByIndex(i);
    for hyperlink in sheet.Hyperlinks do
    begin
      SplitHyperlink(hyperlink^.Target, target, ABookmark);
      if (target <> '') or (ABookmark = '') then
        continue;
      if ParseSheetCellString(ABookmark, sh, r, c) then
        if (sh = TsWorksheet(ACell^.Worksheet).Name) and
           (r = ACell^.Row) and (c = ACell^.Col)
        then
          exit(true);
      if (sheet = FWorksheet) and  ParseCellString(ABookmark, r, c) then
        if (r = ACell^.Row) and (c = ACell^.Col) then
          exit(true);
    end;
  end;
end;

procedure TsHTMLWriter.WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  // nothing to do
end;

procedure TsHTMLWriter.WriteBody(AStream: TStream);
var
  i: Integer;
begin
  AppendToStream(AStream,
    '<body>');
  if HTMLParams.SheetIndex < 0 then      // active sheet
  begin
    if FWorkbook.ActiveWorksheet = nil then
      FWorkbook.SelectWorksheet(FWorkbook.GetWorksheetByIndex(0));
    WriteWorksheet(AStream, FWorkbook.ActiveWorksheet)
  end else
  if HTMLParams.SheetIndex = MaxInt then  // all sheets
    for i:=0 to FWorkbook.GetWorksheetCount-1 do
      WriteWorksheet(AStream, FWorkbook.GetWorksheetByIndex(i))
  else                                    // specific sheet
    WriteWorksheet(AStream, FWorkbook.GetWorksheetbyIndex(HTMLParams.SheetIndex));
  AppendToStream(AStream,
    '</body>');
end;

{ Write boolean cell to stream formatted as string }
procedure TsHTMLWriter.WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: Boolean; ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  AppendToStream(AStream,
    '<div>' + StrUtils.IfThen(AValue, HTMLParams.TrueText, HTMLParams.FalseText) + '</div>');
end;

{ Write date/time values in the same way they are displayed in the sheet }
procedure TsHTMLWriter.WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: TDateTime; ACell: PCell);
var
  s: String;
begin
  Unused(AValue);
  s := FWorksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  s: String;
begin
  Unused(AValue);
  s := FWOrksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

{ HTML does not support formulas, but we can write the formula results to
  to stream. }
procedure TsHTMLWriter.WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  if ACell = nil then
    exit;
  case ACell^.ContentType of
    cctBool      : WriteBool(AStream, ARow, ACol, ACell^.BoolValue, ACell);
    cctEmpty     : ;
    cctDateTime  : WriteDateTime(AStream, ARow, ACol, ACell^.DateTimeValue, ACell);
    cctNumber    : WriteNumber(AStream, ARow, ACol, ACell^.NumberValue, ACell);
    cctUTF8String: WriteLabel(AStream, ARow, ACol, ACell^.UTF8StringValue, ACell);
    cctError     : ;
  end;
end;

{ Writes a LABEL cell to the stream. }
procedure TsHTMLWriter.WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: string; ACell: PCell);
const
  ESCAPEMENT_TAG: Array[TsFontPosition] of String = ('', 'sup', 'sub');
var
  style: String;
  i, n, len: Integer;
  txt, textp, target, bookmark: String;
  rtParam: TsRichTextParam;
  fnt, cellfnt: TsFont;
  escapement: String;
  hyperlink: PsHyperlink;
  isTargetCell: Boolean;
  u: TUri;
begin
  Unused(ARow, ACol, AValue);

  txt := ACell^.UTF8StringValue;
  if txt = '' then
    exit;

  style := ''; //CellFormatAsString(ACell, 'div');
  cellfnt := FWorksheet.ReadCellFont(ACell);

  // Hyperlink
  target := '';
  if FWorksheet.HasHyperlink(ACell) then
  begin
    hyperlink := FWorksheet.FindHyperlink(ACell);
    SplitHyperlink(hyperlink^.Target, target, bookmark);

    n := Length(hyperlink^.Target);
    i := Length(target);
    len := Length(bookmark);

    if (target <> '') and (pos('file:', target) = 0) then
    begin
      u := ParseURI(target);
      if u.Protocol = '' then
        target := '../' + target;
    end;

    // ods absolutely wants "/" path delimiters in the file uri!
    FixHyperlinkPathdelims(target);

    if (bookmark <> '') then
      target := target + '#' + bookmark;
  end;

  // Activate hyperlink target if it is within the same file
  isTargetCell := IsHyperlinkTarget(ACell, bookmark);
  if isTargetCell then bookmark := ' id="' + bookmark + '"' else bookmark := '';

  // No hyperlink, normal text only
  if Length(ACell^.RichTextParams) = 0 then
  begin
    // Standard text formatting
    ValidXMLText(txt);
    if target <> '' then
      txt := Format('<a href="%s">%s</a>', [target, txt]);
    if cellFnt.Position <> fpNormal then
      txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
    AppendToStream(AStream,
      '<div' + bookmark + style + '>' + txt + '</div>')
  end else
  begin
    // "Rich-text" formatted string
    len := UTF8Length(AValue);
    textp := '<div' + bookmark + style + '>';
    if target <> '' then
      textp := textp + '<a href="' + target + '">';
    rtParam := ACell^.RichTextParams[0];
    // Part before first formatted section (has cell fnt)
    if rtParam.StartIndex > 0 then
    begin
      txt := UTF8Copy(AValue, 1, rtParam.StartIndex);
      ValidXMLText(txt);
      if cellfnt.Position <> fpNormal then
        txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
      textp := textp + txt;
    end;
    for i := 0 to High(ACell^.RichTextParams) do
    begin
      // formatted section
      rtParam := ACell^.RichTextParams[i];
      fnt := FWorkbook.GetFont(rtParam.FontIndex);
      style := GetFontAsStyle(rtParam.FontIndex);
      if style <> '' then
        style := ' style="' + style +'"';
      n := rtParam.EndIndex - rtParam.StartIndex;
      txt := UTF8Copy(AValue, rtParam.StartIndex+1, n);
      ValidXMLText(txt);
      if fnt.Position <> fpNormal then
        txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[fnt.Position], txt]);
      textp := textp + '<span' + style +'>' + txt + '</span>';
      // unformatted section before end
      if (rtParam.EndIndex < len) and (i = High(ACell^.RichTextParams)) then
      begin
        txt := UTF8Copy(AValue, rtParam.EndIndex+1, MaxInt);
        ValidXMLText(txt);
        if cellFnt.Position <> fpNormal then
          txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
        textp := textp + txt;
      end else
      // unformatted section between two formatted sections
      if (i < High(ACell^.RichTextParams)) and (rtParam.EndIndex < ACell^.RichTextParams[i+1].StartIndex)
      then begin
        n := ACell^.RichTextParams[i+1].StartIndex - rtParam.EndIndex;
        txt := UTF8Copy(AValue, rtParam.EndIndex+1, n);
        ValidXMLText(txt);
        if cellFnt.Position <> fpNormal then
          txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
        textp := textp + txt;
      end;
    end;
    if target <> '' then
      textp := textp + '</a></div>' else
      textp := textp + '</div>';
    AppendToStream(AStream, textp);
  end;
end;

{ Writes a number cell to the stream. }
procedure TsHTMLWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
var
  s: String;
begin
  Unused(ARow, ACol, AValue);
  s := FWorksheet.ReadAsUTF8Text(ACell, FWorkbook.FormatSettings);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteToStream(AStream: TStream);
begin
  FWorkbook.UpdateCaches;
  AppendToStream(AStream,
    '<!DOCTYPE html>' +
    '<html>' +
      '<head>'+
        '<meta charset="utf-8">');
  WriteStyles(AStream);
  AppendToStream(AStream,
      '</head>');
      WriteBody(AStream);
  AppendToStream(AStream,
    '</html>');
end;

procedure TsHTMLWriter.WriteStyles(AStream: TStream);
var
  i: Integer;
  fmt: PsCellFormat;
  fmtStr: String;
begin
  AppendToStream(AStream,
    '<style>');
  for i:=0 to FWorkbook.GetNumCellFormats-1 do begin
    fmt := FWorkbook.GetPointerToCellFormat(i);
    fmtStr := CellFormatAsString(fmt, 'td');
    if fmtStr <> '' then
      fmtStr := Format('td.style%d {%s}', [i+1, fmtStr]);
    AppendToStream(AStream, fmtStr);
  end;
  AppendToStream(AStream,
    '</style>');
end;

procedure TsHTMLWriter.WriteToStrings(AStrings: TStrings);
var
  Stream: TStream;
begin
  Stream := TStringStream.Create('');
  try
    WriteToStream(Stream);
    Stream.Position := 0;
    AStrings.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TsHTMLWriter.WriteWorksheet(AStream: TStream; ASheet: TsWorksheet);
var
  r, rFirst, rLast: Cardinal;
  c, cFirst, cLast: Cardinal;
  cell: PCell;
  style, s: String;
  fixedLayout: Boolean;
  col: PCol;
  fmt: PsCellFormat;
begin
  FWorksheet := ASheet;

  rFirst := FWorksheet.GetFirstRowIndex;
  cFirst := FWorksheet.GetFirstColIndex;
  rLast := FWorksheet.GetLastOccupiedRowIndex;
  cLast := FWorksheet.GetLastOccupiedColIndex;

  fixedLayout := false;
  for c:=cFirst to cLast do
  begin
    col := FWorksheet.GetCol(c);
    if col <> nil then
    begin
      fixedLayout := true;
      break;
    end;
  end;

  style := GetFontAsStyle(DEFAULT_FONTINDEX);

  style := style + 'border-collapse:collapse; ';
  if soShowGridLines in FWorksheet.Options then
    style := style + GetGridBorderAsStyle;

  if fixedLayout then
    style := style + 'table-layout:fixed; '
  else
    style := style + 'table-layout:auto; width:100%; ';

  AppendToStream(AStream,
    '<div>' +
      '<table style="' + style + '">');

  if HTMLParams.ShowRowColHeaders then
  begin
    // width of row-header column
    style := '';
    if soShowGridLines in FWorksheet.Options then
      style := style + GetGridBorderAsStyle;
    if style <> '' then
      style := ' style="' + style + '"';
    style := style + GetColWidthAsAttr(-1);
    AppendToStream(AStream,
        '<th' + style + '/>');
    // Column headers
    for c := cFirst to cLast do
    begin
      style := '';
      if soShowGridLines in FWorksheet.Options then
        style := style + GetGridBorderAsStyle;
      if style <> '' then
        style := ' style="' + style + '"';
      if fixedLayout then
        style := style + GetColWidthAsAttr(c);
      AppendToStream(AStream,
        '<th' + style + '>' + GetColString(c) + '</th>');
    end;
  end;

  for r := rFirst to rLast do begin
    AppendToStream(AStream,
        '<tr>');

    // Row headers
    if HTMLParams.ShowRowColHeaders then begin
      style := '';
      if soShowGridLines in FWorksheet.Options then
        style := style + GetGridBorderAsStyle;
      if style <> '' then
        style := ' style="' + style + '"';
      style := style + GetRowHeightAsAttr(r);
      AppendToStream(AStream,
          '<th' + style + '>' + IntToStr(r+1) + '</th>');
    end;

    for c := cFirst to cLast do begin
      // Pointer to current cell in loop
      cell := FWorksheet.FindCell(r, c);

      // Cell formatting via predefined styles ("class")
      style := '';
      fmt := nil;
      if cell <> nil then
      begin
        style := Format(' class="style%d"', [cell^.FormatIndex+1]);
        fmt := FWorkbook.GetPointerToCellFormat(cell^.FormatIndex);
      end;

      // Overriding differences between html and fps formatting
      s := '';
      if (fmt = nil) then
        s := s + GetGridBorderAsStyle
      else begin
        if ((not (uffBorder in fmt^.UsedFormattingFields)) or (fmt^.Border = [])) then
          s := s + GetGridBorderAsStyle;
        if ((not (uffHorAlign in fmt^.UsedFormattingFields)) or (fmt^.HorAlignment = haDefault)) then
          s := s + GetDefaultHorAlignAsStyle(cell);
        if ((not (uffVertAlign in fmt^.UsedFormattingFields)) or (fmt^.VertAlignment = vaDefault)) then
          s := s + GetVertAlignAsStyle(vaBottom);
      end;
      if s <> '' then
        style := style + ' style="' + s + '"';

      if not HTMLParams.ShowRowColHeaders then
      begin
        // Column width
        if fixedLayout then
          style := GetColWidthAsAttr(c) + style;

        // Row heights (should be in "tr", but does not work there)
        style := GetRowHeightAsAttr(r) + style;
      end;

      // Merged cells
      if FWorksheet.IsMerged(cell) then
      begin
        if FWorksheet.IsMergeBase(cell) then
          style := style + GetMergedRangeAsStyle(cell)
        else
          Continue;
      end;

      if (cell = nil) or (cell^.ContentType = cctEmpty) then
        // Empty cell
        AppendToStream(AStream,
          '<td' + style + ' />')
      else
      begin
        // Cell with data
        AppendToStream(AStream,
          '<td' + style + '>');
        WriteCellToStream(AStream, cell);
        AppendToStream(AStream,
          '</td>');
      end;
    end;
    AppendToStream(AStream,
        '</tr>');
  end;
  AppendToStream(AStream,
      '</table>' +
    '</div>');
end;

initialization
  InitFormatSettings(HTMLParams.FormatSettings);
  RegisterSpreadFormat(TsHTMLReader, TsHTMLWriter, sfHTML);

end.

