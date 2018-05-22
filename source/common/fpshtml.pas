unit fpsHTML;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fasthtmlparser,
  fpstypes, fpsClasses, fpsReaderWriter, fpsHTMLUtils;

type
  TsHTMLReader = class(TsCustomSpreadReader)
  private
    FPointSeparatorSettings: TFormatSettings;
    FFormatSettings: TFormatSettings;
    parser: THTMLParser;
    FInTable: Boolean;
    FInCell: Boolean;
    FEncoding: String;
    FTableCounter: Integer;
    FCurrRow, FCurrCol: LongInt;
    FCurrCellFormat: TsCellFormat;
    FCurrRichTextParams: TsRichTextParams;
    FCellFont: TsFont;
    FCurrFont: TsFont;
    FCellText: String;
    FAttrList: TsHTMLAttrList;
    FColSpan, FRowSpan: Integer;
    FHRef: String;
    FFontStack: TsIntegerStack;
    FWindowsClipboardMode: Boolean;
    procedure ReadBackgroundColor;
    procedure ReadBorder;
    procedure ReadEncoding;
    procedure ReadFont(AFont: TsFont);
    procedure ReadHRef;
    procedure ReadHorAlign;
    procedure ReadMergedRange;
    procedure ReadTextRot;
    procedure ReadVertAlign;
    procedure ReadWordwrap;
    procedure InitFont(AFont: TsFont);
    procedure InitCellFormat;
    procedure ProcessCellTags(NoCaseTag, Actualtag: String);
    procedure ProcessEndTags(NoCaseTag, ActualTag: String);
    procedure ProcessFontPosition(AFontPosition: TsFontPosition);
    procedure ProcessFontSizeAndStyle(AFontSize: Integer; AFontStyle: TsFontStyles);
    procedure ProcessFontStyle(AFontStyle: TsFontStyle);
    procedure ProcessFontRestore;
    procedure TagFoundHandler(NoCaseTag, ActualTag: string);
    procedure TextFoundHandler(AText: String);
  protected
    procedure AddCell(ARow, ACol: LongInt; AText: String);
    function AddFont(AFont: TsFont): Integer;
    procedure AddRichTextParam(AFont: TsFont; AHyperlinkIndex: Integer = -1);
    procedure FixRichTextParams(var AParams: TsRichTextParams);
  public
    constructor Create(AWorkbook: TsBasicWorkbook); override;
    destructor Destroy; override;
    procedure ReadFromStream(AStream: TStream; APassword: String = '';
      AParams: TsStreamParams = []); override;
    procedure ReadFromStrings(AStrings: TStrings; AParams: TsStreamParams = []); override;
  end;

  TsHTMLWriter = class(TsCustomSpreadWriter)
  private
    FPointSeparatorSettings: TFormatSettings;
    FWindowsClipboardMode: Boolean;
    FStartHtmlPos: Int64;
    FEndHtmlPos: Int64;
    FStartFragmentPos: Int64;
    FEndFragmentPos: Int64;
    function CellFormatAsString(AFormat: PsCellFormat): String;
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
    procedure WriteWorksheet(AStream: TStream; ASheet: TsBasicWorksheet);

  protected
    procedure InternalWriteToStream(AStream: TStream);
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
    constructor Create(AWorkbook: TsBasicWorkbook); override;
    destructor Destroy; override;
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); override;
    procedure WriteToStrings(AStrings: TStrings; AParams: TsStreamParams = []); override;
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
  {@@ Default settings for reading/writing of HTML files }
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

  sfidHTML: TsSpreadFormatID;

implementation

uses
  LConvEncoding, LazUTF8, URIParser, StrUtils, Math,
  fpsUtils, fpspreadsheet, fpsXMLCommon, fpsNumFormat;

const
  MIN_FONTSIZE = 6;

  NATIVE_HEADER  = 'Version:0.9' + #13#10 +
                   'StartHTML:%.10d' + #13#10 +       // Index of first char of <HTML> tag
                   'EndHTML:%.10d' + #13#10 +         // End of end of file
                   'StartFragment:%.10d' + #13#10 +   // Index of first char after <TABLE> tag
                   'EndFragment:%.10d' + #13#10;      // Index of last char before </TABLE> tag

  START_FRAGMENT = '<!--StartFragment-->';

  END_FRAGMENT   = '<!--EndFragment-->';

{==============================================================================}
{                             TsHTMLReader                                     }
{==============================================================================}

constructor TsHTMLReader.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited Create(AWorkbook);
  FEncoding := EncodingUTF8;

  FFormatSettings := HTMLParams.FormatSettings;
  ReplaceFormatSettings(FFormatSettings, FWorkbook.FormatSettings);

  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  FTableCounter := -1;
  FAttrList := TsHTMLAttrList.Create;
  FCellFont := TsFont.Create;
  FCurrFont := TsFont.Create;
  InitFont(FCurrFont);

  FFontStack := TsIntegerStack.Create;
end;

destructor TsHTMLReader.Destroy;
begin
  FreeAndNil(FFontStack);
  FreeAndNil(FCurrFont);
  FreeAndNil(FCellFont);
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
  fntIndex: Integer;
  sheet: TsWorksheet;
begin
  // Empty strings are blank cells -- nothing to do
  if (AText = '') then
    exit;

  sheet := FWorksheet as TsWorksheet;

  // Create cell
  cell := sheet.AddCell(ARow, ACol);

  // Format, rich-text formatting parameters
  // Reject non-used runs; adapt font index to the workbook.
  FixRichTextParams(FCurrRichTextParams);
  // There is only one formatting run which extends across the entire cell
  // --> replace the cell font by that of the formatting run and ignore the run.
  if (Length(FCurrRichTextParams) = 1) and (FCurrRichTextParams[0].FirstIndex = 1) then
  begin
    FCurrCellFormat.FontIndex := FCurrRichTextParams[0].FontIndex;
    SetLength(FCurrRichTextParams, 0);
  end else
  begin
    // Get cell font and use it in the cell format
    fntIndex := (FWorkbook as TsWorkbook).FindFont(FCellFont.FontName, FCellFont.Size,
      FCellFont.Style, FCellFont.Color, FCellFont.Position);
    if fntIndex = -1 then
      fntIndex := (FWorkbook as TsWorkbook).AddFont(FCellFont.FontName, FCellFont.Size,
        FCellFont.Style, FCellFont.Color, FCellFont.Position);
    FCurrCellFormat.FontIndex := fntIndex;
  end;
  if FCurrCellFormat.FontIndex > 0 then
    Include(FCurrCellFormat.UsedFormattingFields, uffFont) else
    Exclude(FCurrCellFormat.UsedFormattingFields, uffFont);
  // Store the cell format in the workbook
  cell^.FormatIndex := (FWorkbook as TsWorkbook).AddCellFormat(FCurrCellFormat);

  // Merged cells
  if (FColSpan > 0) or (FRowSpan > 0) then begin
    sheet.MergeCells(ARow, ACol, ARow + FRowSpan, ACol + FColSpan);
    FRowSpan := 0;
    FColSpan := 0;
  end;

  // Hyperlink
  if FHRef <> '' then begin
    sheet.WriteHyperlink(cell, FHRef);
    FHRef := '';
  end;

  // Case: Do not try to interpret the strings. --> everything is a LABEL cell.
  if not HTMLParams.DetectContentType then
  begin
    sheet.WriteText(cell, AText, FCurrRichTextParams);
    exit;
  end;

  // Check for a NUMBER or CURRENCY cell
  if IsNumberValue(AText, HTMLParams.AutoDetectNumberFormat, FFormatSettings,
    dblValue, nf, decs, currSym, warning) then
  begin
    if currSym <> '' then
      sheet.WriteCurrency(cell, dblValue, nfCurrency, decs, currSym)
    else
      sheet.WriteNumber(cell, dblValue, nf, decs);
    if warning <> '' then
      FWorkbook.AddErrorMsg('Cell %s: %s', [GetCellString(ARow, ACol), warning]);
    exit;
  end;

  // Check for a DATE/TIME cell
  // No idea how to apply the date/time formatsettings here...
  if IsDateTimevalue(AText, FFormatSettings, dtValue, nf) then
  begin
    sheet.WriteDateTime(cell, dtValue, nf);
    exit;
  end;

  // Check for a BOOLEAN cell
  if IsBoolValue(AText, HTMLParams.TrueText, HTMLParams.FalseText, boolValue) then
  begin
    sheet.WriteBoolValue(cell, boolValue);
    exit;
  end;

  // What is left is handled as a TEXT cell
  sheet.WriteText(cell, AText, FCurrRichTextParams);
end;

{ Stores a font in the internal font list. Does not allow duplicates. }
function TsHTMLReader.AddFont(AFont: TsFont): Integer;
const
  EPS = 1e-3;
var
  fnt: TsFont;
begin
  // Is the font already stored in the internal font list?
  for Result := 0 to FFontList.Count-1 do
  begin
    fnt := TsFont(FFontList.Items[Result]);
    if (fnt <> nil) and
       SameText(AFont.FontName, fnt.FontName) and
       SameValue(AFont.Size, fnt.Size, EPS) and
       (AFont.Style = fnt.Style) and
       (AFont.Color = fnt.Color) and
       (AFont.Position = fnt.Position)
    then
      // Yes. Return the font index.
      exit;
  end;

  // No. Create a new font, add it to the list, and return the new index.
  fnt := TsFont.Create;
  fnt.CopyOf(AFont);
  Result := FFontList.Add(fnt);
end;

procedure TsHTMLReader.AddRichTextParam(AFont: TsFont; AHyperlinkIndex: Integer = -1);
var
  len: Integer;
  fntIndex: Integer;
  n: Integer;
begin
  n := Length(FCurrRichTextParams);
  len := UTF8Length(FCellText);
  fntIndex := AddFont(AFont);
  if (n > 0) and (FCurrRichTextparams[n-1].FirstIndex = len+1) then
  begin
    // Avoid adding another rich-text parameter for the same text location:
    // Update the previous one
    FCurrRichTextParams[n-1].FontIndex := fntIndex;
    FCurrRichTextParams[n-1].HyperlinkIndex := AHyperlinkIndex;
  end else
  begin
    // Add a new rich-text parameter
    SetLength(FCurrRichTextParams, n+1);
    FCurrRichTextParams[n].FirstIndex := len + 1;
    FCurrRichTextParams[n].FontIndex := fntIndex;
    FCurrRichTextParams[n].HyperlinkIndex := AHyperlinkIndex;
  end;
end;

{ Remove "zero-width" rich-text parameters, and retain the parameter added last.
  Replace the font index by the one used in the workbook. }
procedure TsHTMLReader.FixRichTextParams(var AParams: TsRichTextParams);
var
  i, j: Integer;
  rtp, nextrtp: TsRichTextParam;
  fnt: TsFont;
  fntIndex: Integer;
begin
  if Length(AParams) = 0 then
    exit;

  // Remove temporary rich-text parameters which were overwritten by their
  // follower. This should not happen, just in case...
  i := High(AParams) - 1;
  while i >= 0 do
  begin
    nextrtp := AParams[i+1];
    rtp := AParams[i];
    if rtp.FirstIndex = nextrtp.FirstIndex then
    begin
      rtp.FontIndex := nextrtp.FontIndex;
      for j:=i+1 to High(AParams)-1 do
      begin
        AParams[j].FirstIndex := AParams[j+1].FirstIndex;
        AParams[j].FontIndex := AParams[j+1].FontIndex;
        AParams[j].HyperlinkIndex := AParams[j+1].HyperlinkIndex;
      end;
      SetLength(AParams, Length(AParams)-1);
    end;
    dec(i);
  end;

  // Replace the internal list font index by the font index of the workbook.
  for i:=0 to High(FCurrRichTextParams) do
  begin
    fnt := TsFont(FFontList[FCurrRichTextParams[i].FontIndex]);
    fntIndex := (FWorkbook as TsWorkbook).FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
    if fntIndex = -1 then
      fntIndex := (FWorkbook as TsWorkbook).AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
    FCurrRichTextParams[i].FontIndex := fntIndex;
  end;
end;

procedure TsHTMLReader.ProcessCellTags(NoCaseTag, ActualTag: String);
begin
  // Pre-sort to speed up finding the tag
  case NoCaseTag[2] of
    'A': if (pos('<A', NoCaseTag) = 1) then
         begin
           FFontStack.Push(AddFont(FCurrFont));
           FCurrFont.FontName := 'Arial';
           //FCurrFont.Size := 10;  use current size
           FCurrFont.Color := scBlue;
           FCurrFont.Style := [fssUnderline];
           FCurrFont.Position := fpNormal;
           FAttrList.Parse(ActualTag);
           ReadHRef;
           AddRichTextParam(FCurrFont);
         end;
    'B': if (NoCaseTag = '<B>') then
           ProcessFontStyle(fssBold)
         else
         if (NoCaseTag = '<BR>') or (NoCaseTag = '<BR/>') or (pos('<BR ', NoCaseTag) = 1) then
           FCellText := FCellText + FPS_LINE_ENDING;
    'D': if (NoCaseTag = '<DEL>') then
           ProcessFontStyle(fssStrikeout)
         else if pos('<DIV ', NoCaseTag) = 1 then
         begin
           FFontStack.Push(AddFont(FCurrFont));
           FAttrList.Parse(ActualTag);
           ReadBackgroundColor;
           ReadBorder;
           ReadHorAlign;
           ReadTextRot;
           ReadVertAlign;
           ReadWordwrap;
           ReadFont(FCurrFont);
           AddRichTextParam(FCurrFont);
         end;
    'E': if (NoCaseTag = '<EM>') then
           ProcessFontStyle(fssItalic);
    'F': if (pos('<FONT ', NoCaseTag) = 1) then
         begin
           FFontStack.Push(AddFont(FCurrFont));
           FAttrList.Parse(ActualTag);
           ReadFont(FCurrFont);
           AddRichTextparam(FCurrFont);
         end;
    'H': case NoCaseTag[3] of
           '1': ProcessFontSizeAndStyle(16, [fssBold]);     // <H1>
           '2': ProcessFontSizeAndStyle(14, [fssBold]);     // <H2>
           '3': ProcessFontSizeAndStyle(12, [fssBold]);     // <H3>
           '4': ProcessFontSizeAndStyle(12, [fssItalic]);   // <H4>
           '5': ProcessFontSizeAndStyle(10, [fssBold]);     // <H5>
           '6': ProcessFontSizeAndStyle(10, [fssItalic]);   // <H6>
         end;
    'I': case NoCaseTag of
           '<I>'     : ProcessFontStyle(fssItalic);
           '<INS>'   : ProcessFontStyle(fssUnderline);
         end;
    'P': if (NoCaseTag = '<P>') or (pos('<P ', NoCaseTag) = 1) then
         begin
           if FCellText <> '' then
             FCellText := FCellText + FPS_LINE_ENDING;
           FFontStack.Push(AddFont(FCurrFont));
           FAttrList.Parse(ActualTag);
           ReadBackgroundColor;
           ReadBorder;
           ReadHorAlign;
           ReadVertAlign;
           ReadWordwrap;
           ReadTextRot;
           ReadFont(FCurrFont);
           AddRichTextParam(FCurrFont);
         end;
    'S': case NoCaseTag of
           '<STRONG>': ProcessFontStyle(fssBold);
           '<S>'     : ProcessFontStyle(fssStrikeout);
           '<SUB>'   : ProcessFontPosition(fpSubscript);
           '<SUP>'   : ProcessFontPosition(fpSuperscript);
           else
             if (pos('<SPAN ', NoCaseTag) = 1) then
             begin
               FFontStack.Push(AddFont(FCurrFont));
               FAttrList.Parse(ActualTag);
               ReadBackgroundColor;
               ReadBorder;
               ReadHorAlign;
               ReadVertAlign;
               ReadWordwrap;
               ReadTextRot;
               ReadFont(FCurrFont);
               AddRichTextparam(FCurrFont);
             end;
         end;
    'U': if (NoCaseTag = '<U>') then
           ProcessFontStyle(fssUnderline);
  end;
end;

procedure TsHTMLReader.ProcessEndTags(NoCaseTag, ActualTag: String);
var
  fntIndex: Integer;
  sheet: TsWorksheet;
begin
  Unused(ActualTag);
  if not FInTable then exit;

  if (NoCaseTag = '</BODY>') then
    ProcessFontRestore;

  if (NoCaseTag = '</TABLE>') then
  begin
    FInTable := false;
    ProcessFontRestore;
    exit;
  end;

  if not FInCell then exit;

  sheet := FWorksheet as TsWorksheet;

  if (NoCaseTag = '</TD>') or (NoCaseTag = '</TH>') then
  begin
    while sheet.IsMerged(sheet.FindCell(FCurrRow, FCurrCol)) do
      inc(FCurrCol);
    AddCell(FCurrRow, FCurrCol, FCellText);
    FInCell := false;
    fntIndex := FFontStack.Pop;
    if fntIndex <> -1 then
      FCurrFont.CopyOf(TsFont(FFontList[fntIndex]));
    exit;
  end;

  // Pre-sort to speed up finding the tag
  case NoCaseTag[3] of
    'A': if (NoCaseTag = '</A>') then begin
           ProcessFontRestore;
           FCellText := FCellText + ' ';
         end;
    'B': if (NoCaseTag = '</B>') then
           ProcessFontRestore;
    'D': if (NoCaseTag = '</DEL>') or (NoCaseTag = '</DIV>') then
           ProcessFontRestore;
    'E': if (NoCaseTag = '</EM>') then
           ProcessFontRestore;
    'F': if (NoCaseTag = '</FONT>') then
           ProcessFontRestore;
    'H': if (NoCaseTag[4] in ['1'..'9']) then
           ProcessFontRestore;
    'I': if (NoCaseTag = '</I>') or (NoCaseTag = '</INS>') then
           ProcessFontRestore;
    'P': if (NoCaseTag = '</P>') then
         begin
           ProcessFontRestore;
           if FCellText <> '' then FCellText := FCellText + FPS_LINE_ENDING;
         end;
    'S': if (NoCaseTag = '</SUB>') or (NoCaseTag = '</SUP>') or
            (NoCaseTag = '</S>') or (NoCaseTag = '</SPAN>') or
            (NoCaseTag = '</STRONG>')
         then
           ProcessFontRestore;
    'U': if (NoCaseTag = '</U>') then
           ProcessFontRestore;
  end;
end;

procedure TsHTMLReader.ProcessFontPosition(AFontPosition: TsFontPosition);
begin
  FFontStack.Push(AddFont(FCurrFont));
  FCurrFont.Position := AFontPosition;
  AddRichTextParam(FCurrFont);
end;

procedure TsHTMLReader.ProcessFontSizeAndStyle(AFontSize: Integer;
  AFontStyle: TsFontStyles);
begin
  FFontStack.Push(AddFont(FCurrFont));
  FCurrFont.Size := AFontSize;
  FCurrFont.Style := AFontStyle;
  AddRichTextparam(FCurrFont);
end;

procedure TsHTMLReader.ProcessFontStyle(AFontStyle: TsFontStyle);
begin
  FFontStack.Push(AddFont(FCurrFont));
  Include(FCurrFont.Style, AFontStyle);
  AddRichTextParam(FCurrFont);
end;

procedure TsHTMLReader.ProcessFontRestore;
var
  fntIndex: Integer;
begin
  fntIndex := FFontStack.Pop;
  if fntIndex > -1 then
  begin
    FCurrFont.CopyOf(TsFont(FFontList[fntIndex]));
    AddRichTextParam(FCurrFont);
  end;
end;

procedure TsHTMLReader.ReadBackgroundColor;
var
  idx: Integer;
begin
  idx := FAttrList.IndexOfName('bgcolor');             // html tag
  if idx = -1 then
    idx := FAttrList.IndexOfName('background-color');  // value taken from "style"
  if idx > -1 then
  begin
    FCurrCellFormat.Background.BgColor := HTMLColorStrToColor(FAttrList[idx].Value);
    FCurrCellFormat.Background.FgColor := FCurrCellFormat.Background.BgColor;
    FCurrCellFormat.Background.Style := fsSolidFill;  // No other fill styles in html
    Include(FCurrCellFormat.UsedFormattingFields, uffBackground);
  end;
end;

procedure TsHTMLReader.ReadBorder;
var
  idx: Integer;
  value: String;

  procedure ReadBorderAttribs(AValue: String; var ABorderStyle: TsCellBorderStyle);
  var
    L: TStringList;
    w: String;
    style: String;
    color: String;
  begin
    L := TStringList.Create;
    try
      L.StrictDelimiter := true;
      L.Delimiter := ' ';
      L.DelimitedText := AValue;
      w := L[0];
      if L.Count > 1 then style := L[1] else style := '';
      if L.Count > 2 then color := L[2] else color := '';
      if (w = 'thin') or (w = '1px') then
        case style of
          'solid'  : ABorderStyle.LineStyle := lsThin;
          'dashed' : ABorderStyle.LineStyle := lsDashed;
          'dotted' : ABorderStyle.LineStyle := lsDotted;
        end
      else
      if (w = 'medium') then
        case style of
          'solid'  : ABorderStyle.LineStyle := lsMedium;
          'dashed' : ABorderStyle.LineStyle := lsMediumDash;
        end
      else
      if (w = 'thick') and (style = 'solid') then
        ABorderStyle.LineStyle := lsThick
      else
      if (w = 'double') then begin
        ABorderStyle.LineStyle := lsDouble;
        if L.Count > 1 then color := L[1];
      end;
      if color <> '' then
        ABorderStyle.Color := HTMLColorStrToColor(color);
    finally
      L.Free;
    end;
  end;

begin
  idx := FAttrList.IndexOfName('border');
  if idx <> -1 then
  begin
    value := FAttrList[idx].Value;
    ReadBorderAttribs(value, FCurrCellFormat.BorderStyles[cbNorth]);
    FCurrCellFormat.BorderStyles[cbEast] := FCurrCellFormat.BorderStyles[cbNorth];
    FCurrCellFormat.BorderStyles[cbSouth] := FCurrCellFormat.BorderStyles[cbNorth];
    FCurrCellFormat.BorderStyles[cbWest] := FCurrCellFormat.BorderStyles[cbNorth];
    FCurrCellFormat.Border := FCurrCellFormat.Border + [cbNorth, cbSouth, cbEast, cbWest];
    Include(FCurrCellFormat.UsedFormattingFields, uffBorder);
  end;

  idx := FAttrList.IndexOfName('border-left');
  if idx <> -1 then
  begin
    Include(FCurrCellFormat.Border, cbWest);
    value := FAttrList[idx].Value;
    ReadBorderAttribs(value, FCurrCellFormat.BorderStyles[cbWest]);
    Include(FCurrCellFormat.UsedFormattingFields, uffBorder);
  end;

  idx := FAttrList.IndexOfName('border-right');
  if idx <> -1 then
  begin
    Include(FCurrCellFormat.Border, cbEast);
    value := FAttrList[idx].Value;
    ReadBorderAttribs(value, FCurrCellFormat.BorderStyles[cbEast]);
    Include(FCurrCellFormat.UsedFormattingFields, uffBorder);
  end;

  idx := FAttrList.IndexofName('border-top');
  if idx <> -1 then
  begin
    Include(FCurrCellFormat.Border, cbNorth);
    value := FAttrList[idx].Value;
    ReadBorderAttribs(value, FCurrCellFormat.BorderStyles[cbNorth]);
    Include(FCurrCellFormat.UsedFormattingFields, uffBorder);
  end;

  idx := FAttrList.IndexOfName('border-bottom');
  if idx <> -1 then
  begin
    Include(FCurrCellFormat.Border, cbSouth);
    value := FAttrList[idx].Value;
    ReadBorderAttribs(value, FCurrCellFormat.BorderStyles[cbSouth]);
    Include(FCurrCellFormat.UsedFormattingFields, uffBorder);
  end;
end;

procedure TsHTMLReader.ReadEncoding;

  function FoundEncoding(AString: string): Boolean;
  // https://encoding.spec.whatwg.org/#encodings
  begin
    Result := true;
    case AString of
      'utf-8', 'utf8', 'unicode-1-1-utf-8':
        FEncoding := 'utf8';
      'cp1250', 'windows-1250', 'x-cp1250':
        FEncoding := 'cp1250';
      'cp1251', 'windows-1251', 'x-cp1251':
        FEncoding := 'cp1251';
      'ansi_x3.4-1968', 'ascii', 'cp1252', 'cp819', 'csisolatin1', 'ibm819',
      'iso-8859-1', 'iso-ir-100', 'iso8859-1', 'iso88591', 'iso_8859-1',
      'iso_8859-1:1987', 'l1', 'latin1', 'us-ascii', 'windows-1252', 'x-cp1252':
        FEncoding := 'cp1252';
      'cp1253', 'windows-1253', 'x-cp1253':
        FEncoding := 'cp1253';
      'cp1254', 'csisolatin5', 'iso-8859-9', 'iso-ir-148', 'iso8859-9', 'iso88599',
      'iso_8859-9', 'iso_8859-9:1989', 'l5', 'latin5', 'windows-1254', 'x-cp1254':
        FEncoding := 'cp1254';
      'cp1255', 'windows-1255', 'x-cp1255':
        FEncoding := 'cp1255';
      'cp1256', 'windows-1256', 'x-cp1256':
        FEncoding := 'cp1256';
      'cp1257', 'windows-1257', 'x-cp1257':
        FEncoding := 'cp1257';
      'cp1258', 'windows-1258', 'x-cp1258':
        FEncoding := 'cp1258';
      '866', 'cp866', 'csibm866', 'ibm866':
        FEncoding := 'cp866';
      'dos-874', 'iso-8859-11', 'iso8859-11', 'iso885911', 'tis-620', 'windows-874':
        FEncoding := 'cp874';
      'csisolatin2', 'iso-8859-2', 'iso-ir-101', 'iso8859-2', 'iso88592',
      'iso_8859-2', 'iso_8859-2:1987', 'l2', 'latin2':
        FEncoding := 'cpiso88592';
      'csisolatin9', 'iso-8859-15', 'iso8859-15', 'iso885915', 'iso_8859-15', 'l9':
        FEncoding := 'cpiso885915';
      'csmacintosh', 'mac', 'macintosh', 'x-mac-roman':
        FEncoding := 'mactintosh';
      'cskoi8r', 'koi', 'koi8', 'koi8-r', 'koi8_r':
        FEncoding := 'koi8';
      'utf-16be':
        FEncoding := 'ucs2be';
      'utf-16', 'utf-16le':
        FEncoding := 'ucs2le';
      else
        // Above site notes also some asian code pages which are not supported by
        // Lazarus encoding utilities.
        Result := false;
    end;
  end;

var
  idx, p: Integer;
  s: String;
begin
  // HTML 5
  idx := FAttrList.IndexOfName('charset');
  if idx > -1 then begin
    s := Lowercase(FAttrList[idx].Value);
    if (s <> '') and FoundEncoding(s) then exit;
  end;

  // HTML 4
  idx := FAttrList.IndexOfName('http-equiv');
  if idx > -1 then
  begin
    idx := FAttrList.IndexOfName('content');
    if idx > -1 then begin
      s := Lowercase(FAttrList[idx].Value);
      p := pos('charset', s);
      if p > 0 then
      begin
        s := copy(s, p+Length('charset')+1, MaxInt);
        p := pos(';', s);
        if p > 0 then s := copy(s, 1, p-1);
        if (s <> '') and FoundEncoding(s) then exit;
      end;
    end;
  end;
end;

procedure TsHTMLReader.ReadFont(AFont: TsFont);
const
  Factor = 1.2;
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
    defFntSize := (FWorkbook as TsWorkbook).GetDefaultFont.Size;
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

procedure TsHTMLReader.ReadHorAlign;
var
  idx: Integer;
begin
  idx := FAttrList.IndexOfName('align');         // html tag
  if idx = -1 then
    idx := FAttrList.IndexOfName('text-align');  // value taken from "style"
  if idx > -1 then
  begin
    case FAttrList[idx].Value of
      'left'   : FCurrCellFormat.HorAlignment := haLeft;
      'center' : FCurrCellFormat.HorAlignment := haCenter;
      'right'  : FCurrCellFormat.HorAlignment := haRight;
      // -- not implemented in fps
      // 'justify'
      // 'char"
      else      exit;
    end;
    Include(FCurrCellFormat.UsedFormattingFields, uffHorAlign);
  end;
end;

procedure TsHTMLReader.ReadHRef;
var
  idx: Integer;
begin
  FHRef := '';
  idx := FAttrList.IndexOfName('href');
  if idx > -1 then
    FHRef := FAttrList[idx].Value;
end;

procedure TsHTMLReader.ReadMergedRange;
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

procedure TsHTMLReader.ReadTextRot;
begin
{
  // No - text rotation is too complicated...

  idx := FAttrList.IndexOfName('transform');
  if idx = -1 then
    idx := FAttrList.IndexOfName('-moz-transform,');
  if idx = -1 then
    idx := FAttrList.IndexOfName('-o-transform');
  if idx = -1 then
    idx := FAttrList.IndexOfName('-wegkit-transform');
  if idx <> -1 then
  begin
    value := FAttrList[idx].Value;
    p := @value[1];
    while (p <> #0) do begin
      if p^ = '(' then
      begin
        s := '';
        inc(p);
        while (p^ <> #0) and (p^ in ['0'..'9', '.', '+', '-']) do
        begin
          s := s + p^;
          inc(p);
        end;
        break;
      end else
        inc(p);
    end;
    if TryStrToFloat(s, f, FPointSeparatorSettings) then begin
      if f >= 45.0 then
        FCurrCellFormat.TextRotation := rt90DegreeClockwiseRotation
      else if f <= -45 then
        FCurrCellFormat.TextRotation := rt90DegreeCounterClockwiseRotation
    end;
    Include(FCurrCellFormat.UsedFormattingFields, uffTextRotation);
    exit;
  end;

  idx := FAttrList.IndexOfName('text-orientation');
  if idx <> -1 then
    if FAttrList[idx].Value = 'upright' then
    begin
      FCurrCellFormat.TextRotation := rtStacked;
      Include(FCurrCellFormat.UsedFormattingfields, uffTextRotation);
    end;
  }
end;

procedure TsHTMLReader.ReadVertAlign;
var
  idx: Integer;
begin
  idx := FAttrList.IndexOfName('valign');      // html tag
  if idx = -1 then
    idx := FAttrList.IndexOfName('vertical-align');  // style tag
  if idx > -1 then
  begin
    case FAttrList[idx].Value of
      'top'   : FCurrCellFormat.VertAlignment := vaTop;
      'middle': FCurrCellFormat.VertAlignment := vaCenter;
      'bottom': FCurrCellFormat.VertAlignment := vaBottom;
      else      exit;  // others not supported
    end;
    Include(FCurrCellFormat.UsedFormattingFields, uffVertAlign);
  end;
end;

procedure TsHTMLReader.ReadWordwrap;
var
  idx: Integer;
begin
  idx := FAttrList.IndexOfName('word-wrap');
  if idx <> -1 then
  begin
    if FAttrList[idx].Value = 'break-word' then begin
      Include(FCurrCellFormat.UsedFormattingFields, uffWordwrap);
      exit;
    end;
  end;

  idx := FAttrList.IndexOfName('white-space');
  if idx <> -1 then
  begin
    if FAttrList[idx].Value = 'nowrap' then
    begin
      Exclude(FCurrCellFormat.UsedFormattingFields, uffWordwrap);
      exit;
    end;
  end;
end;

procedure TsHTMLReader.InitFont(AFont: TsFont);
var
  fnt: TsFont;
begin
  fnt := (FWorkbook as TsWorkbook).GetDefaultFont;
  AFont.FontName := fnt.FontName;
  AFont.Size := fnt.Size;
  AFont.Style := fnt.Style;
  AFont.Color := fnt.Color;
  AFont.Position := fnt.Position;
end;

procedure TsHTMLReader.InitCellFormat;
begin
  InitFormatRecord(FCurrCellFormat);
  InitFont(FCellFont);

  // HTML tables, by default, have word-wrapped cell texts.
  Include(FCurrCellFormat.UsedFormattingFields, uffWordwrap);

  // Vertical alignment, by default, is "middle"
  FCurrCellFormat.VertAlignment := vaCenter;
  Include(FCurrCellFormat.UsedFormattingFields, uffVertAlign);

  // Clear rich-text parameter list
  SetLength(FCurrRichTextParams, 0);
end;

procedure TsHTMLReader.ReadFromStream(AStream: TStream; APassword: String = '';
  AParams: TsStreamParams = []);
var
  list: TStringList;
begin
  Unused(APassword);
  list := TStringList.Create;
  try
    list.LoadFromStream(AStream);
    ReadFromStrings(list, AParams);
    if (FWorkbook as TsWorkbook).GetWorksheetCount = 0 then
    begin
      FWorkbook.AddErrorMsg('Requested table not found, or no tables in html file');
      TsWorkbook(FWorkbook).AddWorksheet('Dummy');
    end;
  finally
    list.Free;
  end;
end;

procedure TsHTMLReader.ReadFromStrings(AStrings: TStrings;
  AParams: TsStreamParams = []);
begin
  FWindowsClipboardMode := (spWindowsClipboardHTML in AParams);

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
  if (Length(NoCaseTag) > 1) and (NoCaseTag[2] = '/') then
  begin
    ProcessEndTags(NoCaseTag, ActualTag);
    exit;
  end;

  if pos('<META', NoCaseTag) = 1 then
  begin
    FAttrList.Parse(ActualTag);
    ReadEncoding;
    exit;
  end;

  if pos('<BODY ', NoCaseTag) = 1 then
  begin
    InitFont(FCurrFont);
    FAttrList.Parse(ActualTag);
    ReadFont(FCurrFont);
  end else
  if pos('<TABLE', NoCaseTag) = 1 then
  begin
    inc(FTableCounter);
    if (HTMLParams.TableIndex >= 0) and (FTableCounter <> HTMLParams.TableIndex) then
      exit;
    FWorksheet := TsWorkbook(FWorkbook).AddWorksheet(Format('Table #%d', [FTableCounter+1]));
    FInTable := true;
    FCurrRow := -1;
    FCurrCol := -1;
    FFontStack.Push(AddFont(FCurrFont));
    FAttrList.Parse(ActualTag);
    ReadFont(FCurrFont);
    TsWorkbook(FWorkbook).ReplaceFont(DEFAULT_FONTINDEX, FCurrFont.FontName, FCurrFont.Size,
      FCurrFont.Style, FCurrFont.Color, FCurrFont.Position);
    FCellFont.CopyOf(FCurrFont);
    exit;
  end;

  if not FInTable then
    exit;

  // The next tags are processed only within a <TABLE> context
  if (NoCaseTag = '<TR>') or (pos('<TR ', NoCaseTag) = 1) or (NoCaseTag = '<TR/>') then
  begin
    inc(FCurrRow);
    FCurrCol := -1;
  end else
  if (NoCaseTag = '<TD>') or (pos('<TD ', NoCaseTag) = 1) or (NoCaseTag = '<TD/>') or
     (NoCaseTag = '<TH>') or (pos('<TH ', NoCaseTag) = 1) or (NoCaseTag = '<TH/>') then
  begin
    FInCell := true;
    inc(FCurrCol);
    FCellText := '';
    FFontStack.Push(AddFont(FCurrFont));
    InitCellFormat;
    FAttrList.Parse(ActualTag);
    ReadMergedRange;
    ReadBackgroundColor;
    ReadBorder;
    ReadHorAlign;
    ReadTextRot;
    ReadVertAlign;
    ReadWordwrap;
    ReadFont(FCurrFont);
    if NoCaseTag[3] = 'H' then begin        // for <TH>
      Include(FCurrFont.Style, fssBold);
      FCurrCellFormat.HorAlignment := haCenter;
      Include(FCurrCellFormat.UsedFormattingFields, uffHorAlign);
    end;
    FCellFont.CopyOf(FCurrFont);
    exit;
  end;

  if not FInCell then
    exit;

  // The next tags are processed only within a <TD> or <TH> context.
  ProcessCellTags(NoCaseTag, ActualTag);
end;

procedure TsHTMLReader.TextFoundHandler(AText: String);
begin
  if FInCell then
  begin
    AText := CleanHTMLString(ConvertEncoding(AText, FEncoding, EncodingUTF8));
    if AText <> '' then
    begin
      if FCellText = '' then
        FCellText := AText
      else
        FCellText := FCellText + AText;
    end;
  end;
end;

{==============================================================================}
{                             TsHTMLWriter                                     }
{==============================================================================}
constructor TsHTMLWriter.Create(AWorkbook: TsBasicWorkbook);
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

function TsHTMLWriter.CellFormatAsString(AFormat: PsCellFormat): String;
//  ATagName: String): String;
begin
//  Unused(ATagName);
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
  sheet: TsWorksheet;
begin
  sheet := FWorksheet as TsWorksheet;

  if AColIndex < 0 then  // Row header column
  begin
    rLast := sheet.GetLastRowIndex;
    w := (FWorkbook as TsWorkbook).ConvertUnits(Length(IntToStr(rLast)) + 2, suChars, suPoints);
  end else
  begin
    w := sheet.ReadDefaultColWidth(suPoints);
    col := sheet.FindCol(AColIndex);
    if (col <> nil) and (col^.Width > 0) then
      w := (FWorkbook as TsWorkbook).ConvertUnits(col^.Width, FWorkbook.Units, suPoints);
  end;
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
  font := (FWorkbook as TsWorkbook).GetFont(AFontIndex);
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
  (FWorksheet as TsWorksheet).FindMergedRange(AMergeBase, r1, c1, r2, c2);
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
  h := (FWorksheet as TsWorksheet).ReadDefaultRowHeight(suPoints);
  row := (FWorksheet as TsWorksheet).FindRow(ARowIndex);
  if row <> nil then begin
    if row^.RowHeightType = rhtCustom then
      h := abs((FWorkbook as TsWorkbook).ConvertUnits(row^.Height, FWorkbook.Units, suPoints));
  end;
  Result := Format(' height="%.1fpt"', [h], FPointSeparatorSettings);
end;


function TsHTMLWriter.GetTextRotationAsStyle(ATextRot: TsTextRotation): String;
begin
  Unused(ATextRot);
  Result := '';
  (*   --- no - this is not working
  case ATextRot of
    trHorizontal: ;
    rt90DegreeClockwiseRotation:
      Result := 'writing-mode:vertical-rl;transform-origin:left top 0;transform:rotate(90deg);'; //-moz-transform: rotate(90deg);';
//      Result := 'writing-mode:vertical-rl;text-orientation:sideways-right;-moz-transform: rotate(-90deg);';
    rt90DegreeCounterClockwiseRotation:
      Result := 'writing-mode:vertical-rt;transform-origin:left top 0;transform:rotate(-90deg);'; //-moz-transform: rotate(-90deg);';
//    Result := 'writing-mode:vertical-rt;text-orientation:sideways-left;-moz-transform: rotate(-90deg);';
    rtStacked:
      Result := 'writing-mode:vertical-rt;text-orientation:upright;';
  end;
  *)
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

procedure TsHTMLWriter.InternalWriteToStream(AStream: TStream);
begin
  (FWorkbook as TsWorkbook).UpdateCaches;
  AppendToStream(AStream,
    '<!DOCTYPE html>');

  FStartHTMLPos := AStream.Position;

  AppendToStream(AStream,
    '<html>' +
      '<head>'+
        '<meta charset="utf-8">');
  WriteStyles(AStream);
  AppendToStream(AStream,
      '</head>');
      WriteBody(AStream);
  AppendToStream(AStream,
    '</html>');

  FEndHTMLPos := AStream.Position;
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

  for i:=0 to (FWorkbook as TsWorkbook).GetWorksheetCount-1 do
  begin
    sheet := (FWorkbook as TsWorkbook).GetWorksheetByIndex(i);
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
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;

  AppendToStream(AStream,
    '<body>');
  if FWindowsClipboardMode or (HTMLParams.SheetIndex < 0) then      // active sheet
  begin
    if book.ActiveWorksheet = nil then
      book.SelectWorksheet(book.GetWorksheetByIndex(0));
    WriteWorksheet(AStream, book.ActiveWorksheet)
  end else
  if HTMLParams.SheetIndex = MaxInt then  // all sheets
    for i:=0 to book.GetWorksheetCount-1 do
      WriteWorksheet(AStream, book.GetWorksheetByIndex(i))
  else                                    // specific sheet
    WriteWorksheet(AStream, book.GetWorksheetbyIndex(HTMLParams.SheetIndex));
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
  Unused(AValue, ACol, ARow);
  s := (FWorksheet as TsWorksheet).ReadAsText(ACell);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  s: String;
begin
  Unused(AValue, ACol, ARow);
  s := (FWorksheet as TsWorksheet).ReadAsText(ACell);
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
  ESCAPEMENT_TAG: array[TsFontPosition] of String = ('', 'sup', 'sub');
var
  style: String;
  i, n, len: Integer;
  txt, textp, target, bookmark: String;
  rtParam: TsRichTextParam;
  fnt, cellfnt: TsFont;
  hyperlink: PsHyperlink;
  isTargetCell: Boolean;
  u: TUri;
begin
  Unused(ARow, ACol, AValue);

  txt := ACell^.UTF8StringValue;
  if txt = '' then
    exit;

  style := '';
  cellfnt := (FWorksheet as TsWorksheet).ReadCellFont(ACell);

  // Hyperlink
  target := '';
  if FWorksheet.HasHyperlink(ACell) then
  begin
    hyperlink := (FWorksheet as TsWorksheet).FindHyperlink(ACell);
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
    txt := LineEndingToBR(txt);
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
    if rtParam.FirstIndex > 1 then
    begin
      txt := UTF8Copy(AValue, 1, rtParam.FirstIndex - 1);
      ValidXMLText(txt);
      if cellfnt.Position <> fpNormal then
        txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
      textp := textp + txt;
    end;
    for i := 0 to High(ACell^.RichTextParams) do
    begin
      // formatted section
      rtParam := ACell^.RichTextParams[i];
      fnt := (FWorkbook as TsWorkbook).GetFont(rtParam.FontIndex);
      style := GetFontAsStyle(rtParam.FontIndex);
      if style <> '' then
        style := ' style="' + style +'"';
      if i = High(ACell^.RichTextParams) then
        n := len - rtParam.FirstIndex else
        n := ACell^.RichTextParams[i+1].FirstIndex - rtParam.FirstIndex;
      txt := UTF8Copy(AValue, rtParam.FirstIndex, n);
      ValidXMLText(txt);
      if fnt.Position <> fpNormal then
        txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[fnt.Position], txt]);
      textp := textp + '<span' + style +'>' + txt + '</span>';
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
  s := (FWorksheet as TsWorksheet).ReadAsText(ACell, FWorkbook.FormatSettings);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteStyles(AStream: TStream);
var
  i: Integer;
  fmt: PsCellFormat;
  fmtStr: String;
begin
  AppendToStream(AStream,
    '<style>' + LineEnding);
  for i:=0 to (FWorkbook as TsWorkbook).GetNumCellFormats-1 do begin
    fmt := (FWorkbook as TsWorkbook).GetPointerToCellFormat(i);
    fmtStr := CellFormatAsString(fmt);
    if fmtStr <> '' then
      fmtStr := Format('  td.style%d {%s}' + LineEnding, [i+1, fmtStr]);
    AppendToStream(AStream, fmtStr);
  end;
  AppendToStream(AStream,
      'th {background-color:#EFEFEF;text-align:center;}');
  AppendToStream(AStream,
    '</style>' + LineEnding);
end;

procedure TsHTMLWriter.WriteToStream(AStream: TStream; AParams: TsStreamParams = []);
begin
  FWindowsClipboardMode := (spWindowsClipboardHTML in AParams);

  if FWindowsClipboardMode then
  begin
    AppendToStream(AStream, Format(
      NATIVE_HEADER, [0, 0, 0, 0]));  // value will be replaced at end
    InternalWriteToStream(AStream);
    AStream.Position := 0;
    AppendToStream(AStream, Format(
      NATIVE_HEADER, [FStartHTMLPos, FEndHTMLPos, FStartFragmentPos, FEndFragmentPos]));
  end else
    InternalWriteToStream(AStream);
end;

procedure TsHTMLWriter.WriteToStrings(AStrings: TStrings;
  AParams: TsStreamParams = []);
var
  Stream: TStream;
begin
  Stream := TStringStream.Create('');
  try
    WriteToStream(Stream, AParams);
    Stream.Position := 0;
    AStrings.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TsHTMLWriter.WriteWorksheet(AStream: TStream; ASheet: TsBasicWorksheet);
var
  r, rFirst, rLast: LongInt;
  c, cFirst, cLast: LongInt;
  cell: PCell;
  col: PCol;
  row: PRow;
  style, s: String;
  fixedLayout: Boolean;
  fmt: PsCellFormat;
  sheet: TsWorksheet absolute ASheet;
begin
  FWorksheet := ASheet;

  rFirst := sheet.GetFirstRowIndex;
  cFirst := sheet.GetFirstColIndex;
  rLast := sheet.GetLastOccupiedRowIndex;
  cLast := sheet.GetLastOccupiedColIndex;

  fixedLayout := false;
  for c:=cFirst to cLast do
  begin
    col := sheet.GetCol(c);
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
    '<div>' + LineEnding +
      '<table style="' + style + '">' + LineEnding);

  if FWindowsClipboardMode then
  begin
    AppendToStream(AStream, START_FRAGMENT);
    FStartFragmentPos := AStream.Position;
  end;

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
        '  <th' + style + '/>' + LineEnding);
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
      col := sheet.FindCol(c);
      if (col <> nil) and (col^.FormatIndex > 0) then
        style := style + Format(' class="style%d"', [col^.FormatIndex+1]);

      AppendToStream(AStream,
        '  <th' + style + '>' + GetColString(c) + '</th>' + LineEnding);
    end;
  end;

  for r := rFirst to rLast do begin
    row := sheet.FindRow(r);
    AppendToStream(AStream,
        '<tr>' + LineEnding);

    // Row headers
    if HTMLParams.ShowRowColHeaders then begin
      style := '';
      if soShowGridLines in FWorksheet.Options then
        style := style + GetGridBorderAsStyle;
      if style <> '' then
        style := ' style="' + style + '"';
      style := style + GetRowHeightAsAttr(r);
      AppendToStream(AStream,
          '  <th' + style + '>' + IntToStr(r+1) + '</th>' + LineEnding);
    end;

    for c := cFirst to cLast do begin
      // Pointer to current cell in loop
      cell := sheet.FindCell(r, c);
      col := sheet.FindCol(c);

      // Cell formatting via predefined styles ("class")
      style := '';
      fmt := nil;
      if cell <> nil then
      begin
        style := Format(' class="style%d"', [cell^.FormatIndex+1]);
        fmt := (FWorkbook as TsWorkbook).GetPointerToCellFormat(cell^.FormatIndex);
      end else
      if (row <> nil) and (row^.FormatIndex > 0) then
      begin
        style := Format(' class="style%d"', [row^.FormatIndex+1]);
        fmt := (FWorkbook as TsWorkbook).GetPointerToCellFormat(row^.FormatIndex);
      end else
      if (col <> nil) and (col^.FormatIndex > 0) then
      begin
        style := Format(' class="style%d"', [col^.FormatIndex+1]);
        fmt := (FWorkbook as TsWorkbook).GetPointerToCellFormat(col^.FormatIndex);
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
      if sheet.IsMerged(cell) then
      begin
        if sheet.IsMergeBase(cell) then
          style := style + GetMergedRangeAsStyle(cell)
        else
          Continue;
      end;

      if (cell = nil) or (cell^.ContentType = cctEmpty) then
        // Empty cell
        AppendToStream(AStream,
          '  <td' + style + ' />' + LineEnding)
      else
      begin
        // Cell with data
        AppendToStream(AStream,
          '  <td' + style + '>');
        WriteCellToStream(AStream, cell);
        AppendToStream(AStream,
          '</td>' + LineEnding);
      end;
    end;
    AppendToStream(AStream,
        '</tr>' + LineEnding);
  end;

  if FWindowsClipboardMode then
  begin
    AppendToStream(AStream, END_FRAGMENT);
    FEndFragmentPos := AStream.Position;
  end;

  AppendToStream(AStream,
      '</table>' + LineEnding +
    '</div>');
end;


initialization
  InitFormatSettings(HTMLParams.FormatSettings);

  // Registers this reader / writer in fpSpreadsheet
  sfidHTML := RegisterSpreadFormat(sfHTML,
    TsHTMLReader, TsHTMLWriter,
    STR_FILEFORMAT_HTML, 'HTML', [STR_HTML_EXTENSION, '.htm']
  );


end.

