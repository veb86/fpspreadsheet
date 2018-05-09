{
xlsbiff8.pas

Writes an Excel 8 file

An Excel worksheet stream consists of a number of subsequent records.
To ensure a properly formed file, the following order must be respected:

1st record:        BOF
2nd to Nth record: Any record
Last record:       EOF

Excel 8 files are OLE compound document files, and must be written using the
fpOLE library.

Records Needed to Make a BIFF8 File Microsoft Excel Can Use:

Required Records:

BOF - Set the 6 byte offset to 0x0005 (workbook globals)
Window1
FONT - At least five of these records must be included
XF - At least 15 Style XF records and 1 Cell XF record must be included
STYLE
BOUNDSHEET - Include one BOUNDSHEET record per worksheet
EOF

BOF - Set the 6 byte offset to 0x0010 (worksheet)
INDEX
DIMENSIONS
WINDOW2
EOF

The row and column numbering in BIFF files is zero-based.

Excel file format specification obtained from:
http://sc.openoffice.org/excelfileformat.pdf

see also:
http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP005199291.aspx

AUTHORS:  Felipe Monteiro de Carvalho
          Jose Mejuto
}
unit xlsbiff8;

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

// The new OLE code is much better, so always use it
{$define USE_NEW_OLE}
{.$define FPSPREADDEBUG} //define to print out debug info to console. Used to be XLSDEBUG;

interface

uses
  Classes, SysUtils, fpcanvas, DateUtils, contnrs, lazutf8,
  fpstypes, fpspreadsheet, fpsrpn, xlscommon,
  {$ifdef USE_NEW_OLE}
  fpolebasic,
  {$else}
  fpolestorage,
  {$endif}
  fpsutils;

type
  TExternBookKind = (ebkExternal, ebkInternal, ebkAddInFunc, ebkDDE_OLE);

  TBIFF8ExternBook = class
    Kind: TExternBookKind;
    DocumentURL: String;
    SheetNames: String;
    function GetSheetName(AIndex: Integer): String;
  end;

  TBIFF8ExternSheet = packed record
    ExternBookIndex: Word;
    FirstSheetIndex: Word;
    LastSheetIndex: Word;
  end;

  { TsSpreadBIFF8Reader }
  TsSpreadBIFF8Reader = class(TsSpreadBIFFReader)
  private
    PendingRecordSize: SizeInt;
    FSharedStringTable: TStringList;
    FCommentList: TObjectList;
    FCommentPending: Boolean;
    FCommentID: Integer;
    FCommentLen: Integer;
    FBiff8ExternBooks: TFPObjectList;
    FBiff8ExternSheets: array of TBiff8ExternSheet;
    function ReadString(const AStream: TStream; const ALength: Word;
      out ARichTextParams: TsRichTextParams): String;
    function ReadUnformattedWideString(const AStream: TStream;
      const ALength: Word): WideString;
    function ReadWideString(const AStream: TStream; const ALength: Word;
      out ARichTextParams: TsRichTextParams): WideString; overload;
    function ReadWideString(const AStream: TStream;
      const AUse8BitLength: Boolean): WideString; overload;
  protected
    procedure PopulatePalette; override;
    procedure ReadBOUNDSHEET(AStream: TStream);
    procedure ReadCONTINUE(const AStream: TStream);
    procedure ReadDEFINEDNAME(const AStream: TStream);
    procedure ReadEXTERNBOOK(const AStream: TStream);
    procedure ReadEXTERNSHEET(const AStream: TStream);
    procedure ReadFONT(const AStream: TStream);
    procedure ReadFORMAT(AStream: TStream); override;
    procedure ReadHeaderFooter(AStream: TStream; AIsHeader: Boolean); override;
    procedure ReadHyperLink(const AStream: TStream);
    procedure ReadHyperlinkToolTip(const AStream: TStream);
    procedure ReadLABEL(AStream: TStream); override;
    procedure ReadLabelSST(const AStream: TStream);
    procedure ReadMergedCells(const AStream: TStream);
    procedure ReadNOTE(const AStream: TStream);
    procedure ReadOBJ(const AStream: TStream);
//    procedure ReadRichString(const AStream: TStream);
    procedure ReadRPNCellAddress(AStream: TStream; out ARow, ACol: Cardinal;
      out AFlags: TsRelFlags); override;
    procedure ReadRPNCellAddressOffset(AStream: TStream;
      out ARowOffset, AColOffset: Integer; out AFlags: TsRelFlags); override;
    procedure ReadRPNCellRangeAddress(AStream: TStream;
      out ARow1, ACol1, ARow2, ACol2: Cardinal; out AFlags: TsRelFlags); override;
    function ReadRPNCellRange3D(AStream: TStream; var ARPNItem: PRPNItem): Boolean; override;
    procedure ReadRPNCellRangeOffset(AStream: TStream;
      out ARow1Offset, ACol1Offset, ARow2Offset, ACol2Offset: Integer;
      out AFlags: TsRelFlags); override;
    procedure ReadRPNSheetIndex(AStream: TStream; out ASheet1, ASheet2: Integer); override;
    procedure ReadRSTRING(AStream: TStream);
    procedure ReadSST(const AStream: TStream);
    function ReadString_8bitLen(AStream: TStream): String; override;
    procedure ReadStringRecord(AStream: TStream); override;
    procedure ReadTXO(const AStream: TStream);
    procedure ReadXF(const AStream: TStream);

  protected
    procedure ReadWorkbookGlobals(AStream: TStream); override;
    procedure ReadWorksheet(AStream: TStream); override;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    procedure ReadFromStream(AStream: TStream;
      APassword: String = ''; AParams: TsStreamParams = []); override;
  end;

  { TsSpreadBIFF8Writer }

  TsSpreadBIFF8Writer = class(TsSpreadBIFFWriter)
  private
    FSharedStringTable: TStringList;
    FNumStrings: DWord;
  private
    procedure BeginCONTINUERecord(AStream: TStream; out ASizePos: Int64);
    procedure FixRecordSize(AStream: TStream; ASizePos: Int64; ASize: Word);
    function WriteStringHelper(AStream: TStream; const AText: RawByteString;
      const ARichTextParams: TsRichTextParams; Is8BitString: Boolean;
      ABytesAvail: Integer; var ATextIndex, ARichIndex: Integer;
      out AComplete: Boolean): Integer;

  protected
    function GetPrintOptions: Word; override;
    function IndexOfSharedString(const AText: String;
      const ARichTextParams: TsRichTextParams): Integer;
    procedure InternalWriteToStream(AStream: TStream);
    procedure PopulatePalette(AWorkbook: TsWorkbook); override;
    procedure PopulateSharedStringTable(AWorkbook: TsWorkbook);

    { Record writing methods }
    procedure WriteBOF(AStream: TStream; ADataType: Word);
    function  WriteBoundsheet(AStream: TStream; AWorksheet: TsWorksheet): Int64;
    procedure WriteComment(AStream: TStream; ACell: PCell); override;
    procedure WriteComments(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteDefinedName(AStream: TStream; AWorksheet: TsWorksheet;
       const AName: String; AIndexToREF: Word); override;
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteEXTERNBOOK(AStream: TStream);
    procedure WriteEXTERNSHEET(AStream: TStream); override;
    procedure WriteFONT(AStream: TStream; AFont: TsFont);
    procedure WriteFonts(AStream: TStream);
    procedure WriteFORMAT(AStream: TStream; ANumFormatStr: String;
      ANumFormatIndex: Integer); override;
    procedure WriteHeaderFooter(AStream: TStream; AIsHeader: Boolean); override;
    procedure WriteHyperlink(AStream: TStream; AHyperlink: PsHyperlink;
      AWorksheet: TsWorksheet);
    procedure WriteHyperlinks(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteHyperlinkToolTip(AStream: TStream; const ARow, ACol: Cardinal;
      const ATooltip: String);
    procedure WriteINDEX(AStream: TStream);
    procedure WriteLABEL(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteMergedCells(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteMSODrawing1(AStream: TStream; ANumShapes: Word; AComment: PsComment);
    procedure WriteMSODrawing2(AStream: TStream; AComment: PsComment; AObjID: Word);
    procedure WriteMSODrawing2_Data(AStream: TStream; AComment: PsComment; AShapeID: Word);
    procedure WriteMSODrawing3(AStream: TStream);
    procedure WriteNOTE(AStream: TStream; AComment: PsComment; AObjID: Word);
    procedure WriteOBJ(AStream: TStream; AObjID: Word);
    function WriteRichTextStream(AStream: TStream; ABuffer: TMemoryStream;
      ABytesAvail: Integer; out ABytesWritten: Integer;
      out AContinueInString: Boolean): Boolean;
    function WriteRPNCellAddress(AStream: TStream; ARow, ACol: Cardinal;
      AFlags: TsRelFlags): word; override;
    function WriteRPNCellAddress3D(AStream: TStream; ASheet, ARow, ACol: Cardinal;
      AFlags: TsRelFlags): Word; override;
    function WriteRPNCellOffset(AStream: TStream; ARowOffset, AColOffset: Integer;
      AFlags: TsRelFlags): Word; override;
    function WriteRPNCellRangeAddress(AStream: TStream; ARow1, ACol1, ARow2, ACol2: Cardinal;
      AFlags: TsRelFlags): Word; override;
    procedure WriteSST(AStream: TStream);
    function WriteString_8bitLen(AStream: TStream; AString: String): Integer; override;
    procedure WriteSTRINGRecord(AStream: TStream; AString: string); override;
    procedure WriteSTYLE(AStream: TStream);
    procedure WriteTXO(AStream: TStream; AComment: PsComment);
    procedure WriteWINDOW2(AStream: TStream; ASheet: TsWorksheet);
    procedure WriteXF(AStream: TStream; AFormatRecord: PsCellFormat;
      XFType_Prot: Byte = 0); override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False; AParams: TsStreamParams = []); override;
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); override;
  end;

  TExcel8Settings = record
    DateMode: TDateMode;
  end;

var
  {@@ Default settings for reading/writing Excel8 files }
  Excel8Settings: TExcel8Settings = (
    DateMode: dm1900;
  );

  {@@ palette of the 64 default BIFF8 colors as "big-endian color" values }
  PALETTE_BIFF8: array[$00..$3F] of TsColor = (
    $000000,  // $00: black            // 8 built-in default colors
    $FFFFFF,  // $01: white
    $FF0000,  // $02: red
    $00FF00,  // $03: green
    $0000FF,  // $04: blue
    $FFFF00,  // $05: yellow
    $FF00FF,  // $06: magenta
    $00FFFF,  // $07: cyan

    $000000,  // $08: EGA black                          1
    $FFFFFF,  // $09: EGA white                          2
    $FF0000,  // $0A: EGA red                            3
    $00FF00,  // $0B: EGA green                          4
    $0000FF,  // $0C: EGA blue                           5
    $FFFF00,  // $0D: EGA yellow                         6
    $FF00FF,  // $0E: EGA magenta                        7      pink
    $00FFFF,  // $0F: EGA cyan                           8      turqoise

    $800000,  // $10=16: EGA dark red                    9
    $008000,  // $11=17: EGA dark green                 10
    $000080,  // $12=18: EGA dark blue                  11
    $808000,  // $13=19: EGA olive                      12      dark yellow
    $800080,  // $14=20: EGA purple                     13      violet
    $008080,  // $15=21: EGA teal                       14
    $C0C0C0,  // $16=22: EGA silver                     15       gray 25%
    $808080,  // $17=23: EGA gray                       16       gray 50%
    $9999FF,  // $18=24: Periwinkle
    $993366,  // $19=25: Plum
    $FFFFCC,  // $1A=26: Ivory
    $CCFFFF,  // $1B=27: Light turquoise
    $660066,  // $1C=28: Dark purple
    $FF8080,  // $1D=29: Coral
    $0066CC,  // $1E=30: Ocean blue
    $CCCCFF,  // $1F=31: Ice blue

    $000080,  // $20=32: Navy (repeated)
    $FF00FF,  // $21=33: Pink (magenta repeated)
    $FFFF00,  // $22=34: Yellow (repeated)
    $00FFFF,  // $23=35: Turqoise (=cyan repeated)
    $800080,  // $24=36: Purple (repeated)
    $800000,  // $25=37: Dark red (repeated)
    $008080,  // $26=38: Teal (repeated)
    $0000FF,  // $27=39: Blue (repeated)
    $00CCFF,  // $28=40: Sky blue
    $CCFFFF,  // $29=41: Light turquoise (repeated)
    $CCFFCC,  // $2A=42: Light green
    $FFFF99,  // $2B=43: Light yellow
    $99CCFF,  // $2C=44: Pale blue
    $FF99CC,  // $2D=45: rose
    $CC99FF,  // $2E=46: lavander
    $FFCC99,  // $2F=47: tan

    $3366FF,  // $30=48: Light blue
    $33CCCC,  // $31=49: Aqua
    $99CC00,  // $32=50: Lime
    $FFCC00,  // $33=51: Gold
    $FF9900,  // $34=52: Light orange
    $FF6600,  // $35=53: Orange
    $666699,  // $36=54: Blue gray
    $969696,  // $37=55: Gray 40%
    $003366,  // $38=56: Dark teal
    $339966,  // $39=57: Sea green
    $003300,  // $3A=58: very dark green
    $333300,  // $3B=59: olive green
    $993300,  // $3C=60: brown
    $993366,  // $3D=61: plum
    $333399,  // $3E=62: indigo
    $333333   // $3F=63: gray 80%
  );
  // color names according to http://dmcritchie.mvps.org/EXCEL/COLORS.HTM

  sfidExcel8: TsSpreadFormatID;

procedure InitBIFF8Limitations(out ALimitations: TsSpreadsheetFormatLimitations);


implementation

uses
  Math, lconvencoding, LazFileUtils, URIParser,
  fpsStrings, {%H-}fpsPatches, fpsStreams, fpsReaderWriter, fpsPalette,
  fpsNumFormat, fpsExprParser, xlsEscher;

const
   { Excel record IDs }
     INT_EXCEL_ID_MERGEDCELLS            = $00E5;  // BIFF8 only
     INT_EXCEL_ID_MSODRAWING             = $00EC;  // BIFF8 only
     INT_EXCEL_ID_SST                    = $00FC;  // BIFF8 only
     INT_EXCEL_ID_LABELSST               = $00FD;  // BIFF8 only
     INT_EXCEL_ID_EXTERNBOOK             = $01AE;  // BIFF8 only
     INT_EXCEL_ID_TXO                    = $01B6;  // BIFF8 only
     INT_EXCEL_ID_HYPERLINK              = $01B8;  // BIFF8 only
     INT_EXCEL_ID_HLINKTOOLTIP           = $0800;  // BIFF8 only
{%H-}INT_EXCEL_ID_FORCEFULLCALCULATION   = $08A3;

   { Excel OBJ subrecord IDs }
     INT_EXCEL_OBJID_FTEND               = $0000;
{%H-}INT_EXCEL_OBJID_FTMACRO             = $0004;
{%H-}INT_EXCEL_OBJID_FTBUTTON            = $0005;
{%H-}INT_EXCEL_OBJID_FTGMO               = $0006;  // Group marker
{%H-}INT_EXCEL_OBJID_CF                  = $0007;  // Clipboard format
{%H-}INT_EXCEL_OBJID_PIOGRBIT            = $0008;  // Picture option flags
{%H-}INT_EXCEL_OBJID_PICTFMLA            = $0009;  // Picture fmla-style macro
{%H-}INT_EXCEL_OBJID_FTCBLS              = $000A;  // Checkbox link
{%H-}INT_EXCEL_OBJID_FTRBO               = $000B;  // Radio button
{%H-}INT_EXCEL_OBJID_FTSBS               = $000C;  // Scrollbar
{%H-}INT_EXCEL_OBJID_FTNTS               = $000D;  // Notes structure (= Comment)
{%H-}INT_EXCEL_OBJID_FTSBSFMLA           = $000E;  // Scroll bar fmla-style macro
{%H-}INT_EXCEL_OBJID_FTGBODATA           = $000F;  // Group box data
{%H-}INT_EXCEL_OBJID_FTEDODATA           = $0010;  // Edit control data
{%H-}INT_EXCEL_OBJID_FTRBODATA           = $0011;  // Radio button data
{%H-}INT_EXCEL_OBJID_FTCBLSDATA          = $0012;  // Check box data
{%H-}INT_EXCEL_OBJID_FTLBSDATA           = $0013;  // List box data
{%H-}INT_EXCEL_OBJID_FTCBLSFMLA          = $0014;  // Check box link fmla-style macro
     INT_EXCEL_OBJID_FTCMO               = $0015;  // Common object data

   { Cell Addresses constants }
     MASK_EXCEL_COL_BITS_BIFF8           = $00FF;
     MASK_EXCEL_RELATIVE_COL_BIFF8       = $4000;  // This is according to Microsoft documentation,
     MASK_EXCEL_RELATIVE_ROW_BIFF8       = $8000;  // but opposite to OpenOffice documentation!

   { BOF record constants }
     INT_BOF_BIFF8_VER                   = $0600;
     INT_BOF_WORKBOOK_GLOBALS            = $0005;
{%H-}INT_BOF_VB_MODULE                   = $0006;
     INT_BOF_SHEET                       = $0010;
{%H-}INT_BOF_CHART                       = $0020;
{%H-}INT_BOF_MACRO_SHEET                 = $0040;
{%H-}INT_BOF_WORKSPACE                   = $0100;
     INT_BOF_BUILD_ID                    = $1FD2;
     INT_BOF_BUILD_YEAR                  = $07CD;

   { STYLE record constants }
     MASK_STYLE_BUILT_IN                 = $8000;

   { XF substructures }

   { XF_ROTATION }
     XF_ROTATION_HORIZONTAL              = 0;
     XF_ROTATION_90DEG_CCW               = 90;
     XF_ROTATION_90DEG_CW                = 180;
     XF_ROTATION_STACKED                 = 255;   // Letters stacked top to bottom, but not rotated

     TEXT_ROTATIONS: Array[TsTextRotation] of Byte = (
       XF_ROTATION_HORIZONTAL,
       XF_ROTATION_90DEG_CW,
       XF_ROTATION_90DEG_CCW,
       XF_ROTATION_STACKED
     );

   { XF CELL BORDER LINE STYLES }
     MASK_XF_BORDER_LEFT                 = $0000000F;
     MASK_XF_BORDER_RIGHT                = $000000F0;
     MASK_XF_BORDER_TOP                  = $00000F00;
     MASK_XF_BORDER_BOTTOM               = $0000F000;
     MASK_XF_BORDER_DIAGONAL             = $01E00000;

     MASK_XF_BORDER_SHOW_DIAGONAL_DOWN   = $40000000;
     MASK_XF_BORDER_SHOW_DIAGONAL_UP     = $80000000;

   { XF CELL BORDER COLORS }
     MASK_XF_BORDER_LEFT_COLOR           = $007F0000;
     MASK_XF_BORDER_RIGHT_COLOR          = $3F800000;
     MASK_XF_BORDER_TOP_COLOR            = $0000007F;
     MASK_XF_BORDER_BOTTOM_COLOR         = $00003F80;
     MASK_XF_BORDER_DIAGONAL_COLOR       = $001FC000;

   { XF CELL BACKGROUND PATTERN }
     MASK_XF_BACKGROUND_PATTERN          = $FC000000;

   { HLINK FLAGS }
     MASK_HLINK_LINK                     = $00000001;
     MASK_HLINK_ABSOLUTE                 = $00000002;
     MASK_HLINK_DESCRIPTION              = $00000014;
     MASK_HLINK_TEXTMARK                 = $00000008;
{%H-}MASK_HLINK_TARGETFRAME              = $00000080;
{%H-}MASK_HLINK_UNCPATH                  = $00000100;

   { RIGHT-TO-LEFT FLAG }
     MASK_XF_BIDI                        = $C0;

     SHAPEID_BASE = 1024;

     MAX_BYTES_IN_RECORD = 8224;   // without header
     MAX_CHARS_IN_WIDESTRING = 32758;


type
  TBIFF8_DimensionsRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FirstRow: DWord;
    LastRowPlus1: DWord;
    FirstCol: Word;
    LastColPlus1: Word;
    NotUsed: Word;
  end;

  TBIFF8_LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    TextLen: Word;
    TextFlags: Byte;
  end;

  TBIFF8_LabelSSTRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    SSTIndex: DWord;
  end;

  TBiff8_RichTextFormattingRun = packed record
    FirstIndex: Word;
    FontIndex: Word;
  end;

  TBiff8_RichTextFormattingRuns = array of TBiff8_RichTextFormattingRun;

  TBIFF8_XFRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FontIndex: Word;
    NumFormatIndex: Word;
    XFType_Prot_ParentXF: Word;
    Align_TextBreak: Byte;
    TextRotation: Byte;
    Indent_Shrink_TextDir: Byte;
    UsedAttrib: Byte;
    Border_BkGr1: DWord;
    Border_BkGr2: DWord;
    BkGr3: Word;
  end;

  TBIFF8TXORecord = packed record
    RecordID: Word;
    RecordSize: Word;
    OptionFlags: Word;
    TextRot: Word;
    Reserved1: Word;
    Reserved2: Word;
    Reserved3: Word;
    TextLen: Word;
    NumFormattingRuns: Word;
    Reserved4: Word;
    Reserved5: Word;
  end;

  TBIFF8Comment = class
    ID: Integer;
    Text: String;
  end;


procedure InitBIFF8Limitations(out ALimitations: TsSpreadsheetFormatLimitations);
begin
  InitBiffLimitations(ALimitations);
end;


{ TBiff8ExternBook }

function TBiff8ExternBook.GetSheetName(AIndex: Integer): String;
var
  L: TStrings;
begin
  L := TStringList.Create;
  try
    L.Delimiter := #1;
    L.StrictDelimiter := true;
    L.DelimitedText := SheetNames;
    if (AIndex >= 0) and (AIndex < L.Count) then
      Result := L[AIndex]
    else
      Result := '';
  finally
    L.Free;
  end;
end;


{ TsSpreadBIFF8Reader }

constructor TsSpreadBIFF8Reader.Create(AWorkbook: TsWorkbook);
begin
  inherited;
  InitBIFF8Limitations(FLimitations);
end;

destructor TsSpreadBIFF8Reader.Destroy;
var
  j: Integer;
begin
  SetLength(FBiff8ExternSheets, 0);
  FBiff8ExternBooks.Free;

  if Assigned(FSharedStringTable) then
  begin
    for j := FSharedStringTable.Count-1 downto 0 do
      if FSharedStringTable.Objects[j] <> nil then
        FSharedStringTable.Objects[j].Free;
    FSharedStringTable.Free;
  end;

  if Assigned(FCommentList) then
    FCommentList.Free;

  inherited;
end;


{@@ ----------------------------------------------------------------------------
  Populates the reader's default palette using the BIFF8 default colors.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.PopulatePalette;
begin
  FPalette.Clear;
  FPalette.UseColors(PALETTE_BIFF8);
end;

{@@ ----------------------------------------------------------------------------
  Reads a CONTINUE record. If the Flag "FCommentPending" is active then this
  record contains the text of a comment assigned to a cell. The length of the
  string is taken from the preceeding TXO record, and the ID of the comment is
  extracted in another preceeding record, an OBJ record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.ReadCONTINUE(const AStream: TStream);
var
  commentStr: String;
  comment: TBIFF8Comment;
  rtParams: TsRichTextParams;
begin
  if FCommentPending and (FCommentLen > 0) then begin
    commentStr := Utf8Encode(ReadWideString(AStream, FCommentLen, rtParams));
    if commentStr <> '' then
    begin
      comment := TBIFF8Comment.Create;
      comment.ID := FCommentID;
      comment.Text := commentStr;
      FCommentList.Add(comment);
    end;
    FCommentPending := false;
  end;
end;

{ Reads a NOTE record (comment associated with a cell). All comments have been
  collected in the FCommentList of the reader from preceding OBJ, TXO and
  CONTINUE records. }
procedure TsSpreadBIFF8Reader.ReadNOTE(const AStream: TStream);
var
  r, c: Word;
  commentID: Word;
  commentText: String;
  i: Integer;
begin
  { Row of the comment }
  r := WordLEToN(AStream.ReadWord);
  { Column of the comment }
  c := WordLEToN(AStream.ReadWord);
  { Option flags, not needed }
  WordLEToN(AStream.ReadWord);
  { Comment ID }
  commentID := WordLEToN(AStream.ReadWord);
  { Next would be the author - ignored... }

  { Seek comment with this ID in the comment list of the reader. }
  for i:=0 to FCommentList.Count-1 do
    if TBIFF8Comment(FCommentList[i]).ID = commentID then
    begin
      commentText := TBIFF8Comment(FCommentList[i]).Text;
      FWorksheet.WriteComment(r, c, commentText);
      exit;
    end;
end;

{ Reads an OBJ record. So far, we only evaluate it to get the ID of the comment
  stored in a following TXO and CONTINUE record. }
procedure TsSpreadBIFF8Reader.ReadOBJ(const AStream: TStream);
var
  subrecID, subrecSize: Word;
  streamPos, p: Int64;
  streamSize: Int64;
  objType: Word;
  objID: Word;
begin
  streamSize := AStream.Size;
  while true do
  begin
    subrecID := WordLEToN(AStream.ReadWord);
    subrecSize := WordLEToN(AStream.ReadWord);
    streamPos := AStream.Position;
    case subrecID of
      INT_EXCEL_OBJID_FTCMO:  // common object data
        // This is the first sub-record of the OBJ record.
        begin
          objType := WordLEToN(AStream.ReadWord);
          objID := WordLEToN(AStream.ReadWord);
          if objType = $19 then begin  // $19 = object is a "comment"
            FCommentPending := true;
            FCommentID := objID;
            FCommentLen := 0;  // will be determined in following TX0 record
            exit;
          end else
            FCommentPending := false;
        end;

      INT_EXCEL_OBJID_FTLBSDATA:
        if subrecSize = $1FEE then   // this cannot be the true sub-record size !!!
          // https://mail-archives.apache.org/mod_mbox/poi-dev/200409.mbox/%3CC1ECA5ECAA06A64D88D955E9E152680D32A5C5@SNOWBALL2.asc.com.au%3E
          // "Every sheet I have looked at seems to have a 16 byte ftLbsData sub-record."
          subrecSize := 16
          // NOTE:
          // This is a risky assumption. A more robust implementation must look at
          // the individual elements of this subrecord, see https://searchcode.com/codesearch/view/47124816/
        else
        if subrecSize = 0 then
        // From MS doc: "If cbFContinued is 0x0000, all of the fields in this
        // structure except ft and cbFContinued MUST NOT exist."
          exit;  // We exit because the stream position cannot advance any more!

      INT_EXCEL_OBJID_FTEND:
        // This is the last sub-record.
        exit;
    end;

    // The structure of the OBJ records is very chaotic. Therefore, it can easily
    // occur that we are lost and read beyond stream end: Check for stream end
    // and store an error. Normal reading will we resumed at the correct position
    // by the main reading loop.
    p := streamPos + subrecSize;
    if p < streamSize then
      AStream.Position := p
    else begin
      FWorkbook.AddErrorMsg(Format(rsFileStructureError, ['OBJ', streamPos]));
      exit;
    end;
  end;
end;

{ Reads a unicode string which does not contain rich-text information.
  This is needed for the RSTRING record. }
function TsSpreadBIFF8Reader.ReadUnformattedWideString(const AStream: TStream;
  const ALength: WORD): WideString;
var
  flags: Byte;
  DecomprStrValue: WideString;
  i: Integer;
  len: SizeInt;
  recType: Word;
  {%H-}recSize: Word;
  C: WideChar;
begin
  flags := AStream.ReadByte;
  dec(PendingRecordSize);
  if flags and 1 = 1 Then begin
    //String is WideStringLE
    if (ALength * SizeOf(WideChar)) > PendingRecordSize then begin
      SetLength(Result, PendingRecordSize div 2);
      AStream.ReadBuffer(Result[1], PendingRecordSize);
      Dec(PendingRecordSize, PendingRecordSize);
    end else begin
      SetLength(Result, ALength);
      AStream.ReadBuffer(Result[1], ALength * SizeOf(WideChar));
      Dec(PendingRecordSize, ALength * SizeOf(WideChar));
    end;
    Result := WideStringLEToN(Result);
  end else begin
    // String is 1 byte per char, this is UTF-16 with the high byte ommited
    // because it is zero, so decompress and then convert
    len := ALength;
    SetLength(DecomprStrValue, len);
    for i := 1 to len do
    begin
      C := WideChar(AStream.ReadByte);  // Read 1 byte, but put it into a 2-byte char
      DecomprStrValue[i] := C;
      dec(PendingRecordSize);
      if (PendingRecordSize <= 0) and (i < len) then begin
        //A CONTINUE may have happened here
        recType := WordLEToN(AStream.ReadWord);
        recSize := WordLEToN(AStream.ReadWord);
        if recType <> INT_EXCEL_ID_CONTINUE then begin
          raise EFPSpreadsheetReader.Create('[TsSpreadBIFF8Reader.ReadWideString] Expected CONTINUE record not found.');
        end else begin
          PendingRecordSize := RecordSize;
          DecomprStrValue := copy(DecomprStrValue,1,i) + ReadUnformattedWideString(AStream, ALength-i);
          break;
        end;
      end;
    end;
    Result := DecomprStrValue;
  end;
end;

function TsSpreadBIFF8Reader.ReadWideString(const AStream: TStream;
  const ALength: WORD; out ARichTextParams: TsRichTextParams): WideString;
var
  StringFlags: BYTE;
  DecomprStrValue: WideString;
  AnsiStrValue: ansistring;
  RunsCounter: WORD;
  AsianPhoneticBytes: DWORD;
  rtf_dummy: TsRichTextParams;
  i: Integer;
  j: Integer; //j: SizeUInt;
  lLen: SizeInt;
  recType: WORD;
  recSize: WORD;
  C: WideChar;
begin
  StringFlags := AStream.ReadByte;
  Dec(PendingRecordSize);
  if StringFlags and 8 = 8 then begin
    // Rich string
    RunsCounter := WordLEtoN(AStream.ReadWord);
    dec(PendingRecordSize,2);
  end;
  if StringFlags and 4 = 4 then begin
    // Asian phonetics
    // Read Asian phonetics Length (not used)
    AsianPhoneticBytes := DWordLEtoN(AStream.ReadDWord);
    dec(PendingRecordSize,4);
  end;
  if StringFlags and 1 = 1 Then begin
    // String is WideStringLE
    if (ALength*SizeOf(WideChar)) > PendingRecordSize then begin
      SetLength(Result, PendingRecordSize div 2);
      AStream.ReadBuffer(Result[1], PendingRecordSize);
      Dec(PendingRecordSize, PendingRecordSize);
      // We reached the end of the record and switch to the CONTINUE record
      recType := WordLEToN(AStream.ReadWord);
      recSize := WordLEToN(AStream.ReadWord);
      if recType <> INT_EXCEL_ID_CONTINUE then
        raise EFPSpreadsheetReader.Create('[TsSpreadBIFF8Reader.ReadWideString] CONTINUE record expected, but not found.');
      PendingRecordSize := recSize;
      Result := Result + ReadWideString(AStream, ALength - Length(Result), rtf_dummy);
    end else begin
      SetLength(Result, ALength);
      AStream.ReadBuffer(Result[1], ALength * SizeOf(WideChar));
      Dec(PendingRecordSize, ALength * SizeOf(WideChar));
    end;
    Result := WideStringLEToN(Result);
  end else begin
    // String is 1 byte per char, this is UTF-16 with the high byte ommited
    // because it is zero, so decompress and then convert
    lLen := ALength;
    SetLength(DecomprStrValue, lLen);
    for i := 1 to lLen do
    begin
      C := WideChar(AStream.ReadByte);  // Read 1 byte, but put it into a 2-byte char
      DecomprStrValue[i] := C;
      Dec(PendingRecordSize);
      if (PendingRecordSize <= 0) and (i < lLen) then begin
        //A CONTINUE may have happened here
        recType := WordLEToN(AStream.ReadWord);
        recSize := WordLEToN(AStream.ReadWord);
        if recType <> INT_EXCEL_ID_CONTINUE then begin
          Raise EFPSpreadsheetReader.Create('[TsSpreadBIFF8Reader.ReadWideString] CONTINUE record expected, but not found.');
        end else begin
          PendingRecordSize := recSize;
          DecomprStrValue := copy(DecomprStrValue,1,i) + ReadWideString(AStream, ALength-i, ARichTextParams);
          break;
        end;
      end;
    end;
    Result := DecomprStrValue;
  end;
  if StringFlags and 8 = 8 then begin
    // Rich string (This only occurs in BIFF8)
    SetLength(ARichTextParams, RunsCounter);
    for j := 0 to SmallInt(RunsCounter) - 1 do begin
      if (PendingRecordSize <= 0) then begin
        // A CONTINUE may happened here
        recType := WordLEToN(AStream.ReadWord);
        recSize := WordLEToN(AStream.ReadWord);
        if recType <> INT_EXCEL_ID_CONTINUE then begin
          Raise EFPSpreadsheetReader.Create('[TsSpreadBIFF8Reader.ReadWideString] CONTINUE record expected, but not found.');
        end else begin
          PendingRecordSize := recSize;
        end;
      end;
      // character start index: 0-based in file, 1-based in fps
      ARichTextParams[j].FirstIndex := WordLEToN(AStream.ReadWord) + 1;
      ARichTextParams[j].FontIndex := WordLEToN(AStream.ReadWord);
      ARichTextParams[j].HyperlinkIndex := -1;
      dec(PendingRecordSize, 2*2);
    end;
  end;
  if StringFlags and 4 = 4 then begin
    // Asian phonetics
    // Read Asian phonetics, discarded as not used.
    SetLength(AnsiStrValue, AsianPhoneticBytes);
    AStream.ReadBuffer(AnsiStrValue[1], AsianPhoneticBytes);
    dec(PendingRecordSize, AsianPhoneticBytes);
  end;
end;

function TsSpreadBIFF8Reader.ReadWideString(const AStream: TStream;
  const AUse8BitLength: Boolean): WideString;
var
  Len: Word;
  rtParams: TsRichTextParams;
begin
  if AUse8BitLength then
    Len := AStream.ReadByte()
  else
    Len := WordLEtoN(AStream.ReadWord());

  Result := ReadWideString(AStream, Len, rtParams);
end;

procedure TsSpreadBIFF8Reader.ReadWorkbookGlobals(AStream: TStream);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
begin
  if FCommentList = nil
    then FCommentList := TObjectList.Create
    else FCommentList.Clear;

  if Assigned(FSharedStringTable) then FreeAndNil(FSharedStringTable);

  while (not SectionEOF) do begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);
    PendingRecordSize := RecordSize;

    CurStreamPos := AStream.Position;

    if RecordType <> INT_EXCEL_ID_CONTINUE then begin
      case RecordType of
       INT_EXCEL_ID_BOF           : ;
       INT_EXCEL_ID_BOUNDSHEET    : ReadBoundSheet(AStream);
       INT_EXCEL_ID_CODEPAGE      : ReadCodepage(AStream);
       INT_EXCEL_ID_DATEMODE      : ReadDateMode(AStream);
       INT_EXCEL_ID_DEFINEDNAME   : ReadDEFINEDNAME(AStream);
       INT_EXCEL_ID_EOF           : SectionEOF := True;
       INT_EXCEL_ID_EXTERNBOOK    : ReadEXTERNBOOK(AStream);
       INT_EXCEL_ID_EXTERNSHEET   : ReadEXTERNSHEET(AStream);
       INT_EXCEL_ID_FONT          : ReadFont(AStream);
       INT_EXCEL_ID_FORMAT        : ReadFormat(AStream);
       INT_EXCEL_ID_PALETTE       : ReadPalette(AStream);
       INT_EXCEL_ID_PASSWORD      : ReadPASSWORD(AStream);
       INT_EXCEL_ID_PROTECT       : ReadPROTECT(AStream);
       INT_EXCEL_ID_SST           : ReadSST(AStream);
       INT_EXCEL_ID_WINDOWPROTECT : ReadWindowProtect(AStream);
       INT_EXCEL_ID_XF            : ReadXF(AStream);
    else
        // nothing
      end;
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then SectionEOF := True;
  end;

  // Convert palette indexes to rgb colors
  FixColors;
end;

procedure TsSpreadBIFF8Reader.ReadWorksheet(AStream: TStream);
var
  SectionEOF: Boolean = False;
  RecordType: Word;
  CurStreamPos: Int64;
//  sheetData: TsSheetData;
begin
  (*
  sheetData := TsSheetData(FSheetList[FCurSheetIndex]);
  FWorksheet := FWorkbook.AddWorksheet(sheetData.Name, true);
  if sheetData.Hidden then
    FWorksheet.Options := FWorksheet.Options + [soHidden];
*)
  FWorksheet := FWorkbook.GetWorksheetByIndex(FCurSheetIndex);
  while (not SectionEOF) do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);
    PendingRecordSize := RecordSize;

// For debugging to find out in which record a crash happens:
//    WriteLn(Format('Stream.Pos: %d, RecordType: $%.04x, RecordSize: %d', [AStream.Position-4, RecordType, RecordSize]));

    CurStreamPos := AStream.Position;

    case RecordType of

    INT_EXCEL_ID_BLANK         : ReadBlank(AStream);
    INT_EXCEL_ID_BOF           : ;
    INT_EXCEL_ID_BOOLERROR     : ReadBool(AStream);
    INT_EXCEL_ID_BOTTOMMARGIN  : ReadMargin(AStream, 3);
    INT_EXCEL_ID_COLINFO       : ReadColInfo(AStream);
    INT_EXCEL_ID_CONTINUE      : ReadCONTINUE(AStream);
    INT_EXCEL_ID_DEFCOLWIDTH   : ReadDefColWidth(AStream);
    INT_EXCEL_ID_DEFINEDNAME   : ReadDefinedName(AStream);
    INT_EXCEL_ID_EOF           : SectionEOF := True;
    INT_EXCEL_ID_FOOTER        : ReadHeaderFooter(AStream, false);
    INT_EXCEL_ID_FORMULA       : ReadFormula(AStream);
    INT_EXCEL_ID_HCENTER       : ReadHCENTER(AStream);
    INT_EXCEL_ID_HEADER        : ReadHeaderFooter(AStream, true);
    INT_EXCEL_ID_HLINKTOOLTIP  : ReadHyperlinkToolTip(AStream);
    INT_EXCEL_ID_HYPERLINK     : ReadHyperlink(AStream);
    INT_EXCEL_ID_LABEL         : ReadLabel(AStream);
    INT_EXCEL_ID_LABELSST      : ReadLabelSST(AStream);
    INT_EXCEL_ID_LEFTMARGIN    : ReadMargin(AStream, 0);
    INT_EXCEL_ID_MERGEDCELLS   : ReadMergedCells(AStream);
    INT_EXCEL_ID_MULBLANK      : ReadMulBlank(AStream);
    INT_EXCEL_ID_MULRK         : ReadMulRKValues(AStream);
    INT_EXCEL_ID_NOTE          : ReadNOTE(AStream);
    INT_EXCEL_ID_NUMBER        : ReadNumber(AStream);
    INT_EXCEL_ID_OBJ           : ReadOBJ(AStream);
    INT_EXCEL_ID_OBJECTPROTECT : ReadObjectProtect(AStream, FWorksheet);
    INT_EXCEL_ID_PAGESETUP     : ReadPageSetup(AStream);
    INT_EXCEL_ID_PANE          : ReadPane(AStream);
    INT_EXCEL_ID_PASSWORD      : ReadPASSWORD(AStream, FWorksheet);
    INT_EXCEL_ID_PRINTGRID     : ReadPrintGridLines(AStream);
    INT_EXCEL_ID_PRINTHEADERS  : ReadPrintHeaders(AStream);
    INT_EXCEL_ID_PROTECT       : ReadPROTECT(AStream, FWorksheet);
    INT_EXCEL_ID_RIGHTMARGIN   : ReadMargin(AStream, 1);
    INT_EXCEL_ID_ROW           : ReadRowInfo(AStream);

    //(RSTRING) This record stores a formatted text cell (Rich-Text).
    // In BIFF8 it is usually replaced by the LABELSST record. Excel still
    // uses this record, if it copies formatted text cells to the clipboard.
    INT_EXCEL_ID_RSTRING       : ReadRSTRING(AStream);

    // (RK) This record represents a cell that contains an RK value
    // (encoded integer or floating-point value). If a floating-point
    // value cannot be encoded to an RK value, a NUMBER record will be written.
    // This record replaces the record INTEGER written in BIFF2.
    INT_EXCEL_ID_RK            : ReadRKValue(AStream);

    INT_EXCEL_ID_SCL           : ReadSCLRecord(AStream);
    INT_EXCEL_ID_SELECTION     : ReadSELECTION(AStream);
    INT_EXCEL_ID_SHAREDFMLA    : ReadSharedFormula(AStream);
    INT_EXCEL_ID_SHEETPR       : ReadSHEETPR(AStream);
    INT_EXCEL_ID_STRING        : ReadStringRecord(AStream);
    INT_EXCEL_ID_TOPMARGIN     : ReadMargin(AStream, 2);
    INT_EXCEL_ID_TXO           : ReadTXO(AStream);
    INT_EXCEL_ID_VCENTER       : ReadVCENTER(AStream);
    INT_EXCEL_ID_WINDOW2       : ReadWindow2(AStream);
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then SectionEOF := True;
  end;

  FixCols(FWorksheet);
  FixRows(FWorksheet);
end;

procedure TsSpreadBIFF8Reader.ReadBoundsheet(AStream: TStream);
var
  len: Byte;
  wideName: WideString;
  rtParams: TsRichTextParams;
  sheet: TsWorksheet;
  sheetstate: Byte;
  sheetdata: TsSheetData;
begin
  { Absolute stream position of the BOF record of the sheet represented
    by this record }
  // Just assume that they are in order
  AStream.ReadDWord();

  { Visibility }
  sheetstate := AStream.ReadByte();    // 0=visible, 1=hidden, 2="very" hidden

  { Sheet type }
  AStream.ReadByte();

  { Sheet name: 8-bit length }
  len := AStream.ReadByte();

  { Read string with flags }
  wideName := ReadWideString(AStream, len, rtParams);

  sheet := FWorkbook.AddWorksheet(UTF8Encode(widename), true);
  if sheetState <> 0 then
    sheet.Options := sheet.Options + [soHidden];
(*
  sheetData := TsSheetData.Create;
  sheetData.Name := UTF8Encode(wideName);
  sheetData.Hidden := sheetState <> 0;
  FSheetList.Add(sheetdata);
  *)
end;

function TsSpreadBIFF8Reader.ReadString(const AStream: TStream;
  const ALength: WORD; out ARichTextParams: TsRichTextParams): String;
begin
  Result := UTF16ToUTF8(ReadWideString(AStream, ALength, ARichTextParams));
end;

procedure TsSpreadBIFF8Reader.ReadFromStream(AStream: TStream;
  APassword: String = ''; AParams: TsStreamParams = []);
var
  OLEStream: TMemoryStream;
  OLEStorage: TOLEStorage;
  OLEDocument: TOLEDocument;
begin
  Unused(APassword, AParams);
  OLEStream := TMemoryStream.Create;
  try
    // Only one stream is necessary for any number of worksheets
    OLEStorage := TOLEStorage.Create;
    try
      OLEDocument.Stream := OLEStream;
      OLEStorage.ReadOLEStream(AStream, OLEDocument, 'Workbook');
      InternalReadFromStream(OLEStream);
    finally
      OLEStorage.Free;
    end;
  finally
    OLEStream.Free;
  end;
end;
(*
procedure TsSpreadBIFF8Reader.ReadFromStream(AStream: TStream);
var
  BIFF8EOF: Boolean;
begin
  { Initializations }
  BIFF8EOF := False;

  FWorksheetNames := TStringList.Create;
  FWorksheetNames.Clear;
  FCurrentWorksheet := 0;

  if FCommentList = nil then FCommentList := TObjectList.Create
    else FCommentList.Clear;

  { Read workbook globals }
  ReadWorkbookGlobals(AStream);

  // Check for the end of the file
  if AStream.Position >= AStream.Size then BIFF8EOF := True;

  { Now read all worksheets }
  while (not BIFF8EOF) do
  begin
    //Safe to not read beyond assigned worksheet names.
    if FCurrentWorksheet > FWorksheetNames.Count-1 then break;

    ReadWorksheet(AStream);

    // Check for the end of the file
    if AStream.Position >= AStream.Size then BIFF8EOF := True;

    // Final preparations
    Inc(FCurrentWorksheet);
    if FCurrentWorksheet = FWorksheetNames.Count then BIFF8EOF := True;
    // It can happen in files written by Office97 that the OLE directory is
    // at the end of the file.
  end;

  { Finalizations }
  FWorksheetNames.Free;
end;
    *)

procedure TsSpreadBIFF8Reader.ReadLABEL(AStream: TStream);
var
  L, i: Word;
  ARow, ACol: Cardinal;
  XF: Word;
  wideStrValue: WideString;
  cell: PCell;
  rtParams: TsRichTextParams;
  fntIndex: Integer;
  fnt: TsFont;
begin
  { BIFF Record data: Row, Column, XF Index }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Byte String with 16-bit size }
  L := WordLEtoN(AStream.ReadWord());

  { Read wide string with flags }
  wideStrValue := ReadWideString(AStream, L, rtParams);

  { Save the data }
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);        // "virtual" cell
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);    // "real" cell

  FWorksheet.WriteText(cell, UTF16ToUTF8(wideStrValue));

  { Add attributes }
  ApplyCellFormatting(cell, XF);

  { Apply rich-text formatting }
  if Length(rtParams) > 0 then begin
    SetLength(cell^.RichTextParams, Length(rtParams));
    for i := 0 to High(rtParams) do
    begin
      // Character index where format starts: 0-based in file, 1-based in fps
      cell^.RichTextParams[i].FirstIndex := rtParams[i].FirstIndex + 1;
      // Font index of new format - need to adjust index!
      fntIndex := rtParams[i].FontIndex;
      fnt := TsFont(FFontList[fntIndex]);
      fntIndex := FWorkbook.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
      if fntIndex = -1 then
        fntIndex := FWorkbook.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
      cell^.RichTextParams[i].FontIndex := fntIndex;
      // Hyperlink index, not used here
      cell^.RichTextParams[i].HyperlinkIndex := -1;
    end;
  end;

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF8Reader.ReadMergedCells(const AStream: TStream);
var
  rng: packed record Row1, Row2, Col1, Col2: Word; end;
  i, n: word;
begin
  rng.Row1 := 0;  // to silence the compiler...

  // Count of merged ranges
  n := WordLEToN(AStream.ReadWord);

  for i:=1 to n do begin
    // Read range
    AStream.ReadBuffer(rng, SizeOf(rng));
    // Transfer cell range to worksheet
    FWorksheet.MergeCells(
      WordLEToN(rng.Row1), WordLEToN(rng.Col1),
      WordLEToN(rng.Row2), WordLEToN(rng.Col2)
    );
  end;
end;
                                                 (*
procedure TsSpreadBIFF8Reader.ReadRichString(const AStream: TStream);
var
  L: Word;
  B: WORD;
  ARow, ACol: Cardinal;
  XF: Word;
  strValue: string;
  cell: PCell;
  rtfRuns: TsRichTextFormattingRuns;
begin
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Byte String with 16-bit size }
  L := WordLEtoN(AStream.ReadWord());
  strValue := ReadString(AStream, L, rtfRuns);

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  { Save the data }
  FWorksheet.WriteUTF8Text(cell, strValue);

  {
  // Read rich-text formatting runs
  B := WordLEtoN(AStream.ReadWord);
  SetLength(rtfRuns, B);
  for L := 0 to B-1 do begin
    rtfRuns[L].FirstIndex := WordLEToN(AStream.ReadWord); // Index of first formatted character
    rtfRuns[L].FontIndex := WordLEToN(AStream.ReadByte);  // Index of font used
  end;
   }
  {Add attributes}
  ApplyCellFormatting(cell, XF);
  ApplyRichTextFormattingRuns(cell, rtfRuns);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;
                                                   *)
{ Reads the cell address used in an RPN formula element. Evaluates the corresponding
  bits to distinguish between absolute and relative addresses.
  Overriding the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellAddress(AStream: TStream;
  out ARow, ACol: Cardinal; out AFlags: TsRelFlags);
var
  c: word;
begin
  // Read row index (2 bytes)
  ARow := WordLEToN(AStream.ReadWord);
  // Read column index; it contains info on absolute/relative address
  c := WordLEToN(AStream.ReadWord);
  // Extract column index
  ACol := c and MASK_EXCEL_COL_BITS_BIFF8;
  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
end;

{ Reads the difference between cell row and column indexed of a cell and
  a reference cell.
  Overrides the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellAddressOffset(AStream: TStream;
  out ARowOffset, AColOffset: Integer; out AFlags: TsRelFlags);
var
  dr: SmallInt;
  dc: ShortInt;
  c: Word;
begin
  // 2 bytes for row offset
  dr := ShortInt(WordLEToN(AStream.ReadWord));
  ARowOffset := dr;

  // 2 bytes for column offset
  c := WordLEToN(AStream.ReadWord);
  dc := ShortInt(Lo(c));
  AColOffset := dc;

  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
end;

{ Reads a cell range address used in an RPN formula element.
  Evaluates the corresponding bits to distinguish between absolute and
  relative addresses.
  Overriding the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellRangeAddress(AStream: TStream;
  out ARow1, ACol1, ARow2, ACol2: Cardinal; out AFlags: TsRelFlags);
var
  c1, c2: word;
begin
  // Read row index of first and last rows (2 bytes, each)
  ARow1 := WordLEToN(AStream.ReadWord);
  ARow2 := WordLEToN(AStream.ReadWord);
  // Read column index of first and last columns; they contain info on
  // absolute/relative address
  c1 := WordLEToN(AStream.ReadWord);
  c2 := WordLEToN(AStream.ReadWord);
  // Extract column index of rist and last columns
  ACol1 := c1 and MASK_EXCEL_COL_BITS_BIFF8;
  ACol2 := c2 and MASK_EXCEL_COL_BITS_BIFF8;
  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c1 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c1 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
  if (c2 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol2);
  if (c2 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow2);
end;

function TsSpreadBIFF8Reader.ReadRPNCellRange3D(AStream: TStream;
  var ARPNItem: PRPNItem): Boolean;
var
  sheetIndex1, sheetIndex2: Integer;
  r1, c1, r2, c2: Cardinal;
  flags: TsRelFlags;
begin
  Result := true;
  ReadRPNSheetIndex(AStream, sheetIndex1, sheetIndex2);
  if (sheetIndex1 = -1) or (sheetIndex2 = -1) then
    exit(False);  // unsupported case

  ReadRPNCellRangeAddress(AStream, r1, c1, r2, c2, flags);
  if r2 = $FFFF then r2 := Cardinal(-1);
  if c2 = $FF then c2 := Cardinal(-1);

  ARPNItem := RPNCellRange3D(
    sheetIndex1, r1, c1,
    sheetIndex2, r2, c2,
    flags, ARPNItem
  );
end;

{ Reads the difference between row and column corner indexes of a cell range
  and a reference cell.
  Overriding the implementation in xlscommon. }
procedure TsSpreadBIFF8Reader.ReadRPNCellRangeOffset(AStream: TStream;
  out ARow1Offset, ACol1Offset, ARow2Offset, ACol2Offset: Integer;
  out AFlags: TsRelFlags);
var
  c1, c2: Word;
begin
  // 2 bytes for offset of first row
  ARow1Offset := ShortInt(WordLEToN(AStream.ReadWord));

  // 2 bytes for offset to last row
  ARow2Offset := ShortInt(WordLEToN(AStream.ReadWord));

  // 2 bytes for offset of first column
  c1 := WordLEToN(AStream.ReadWord);
  ACol1Offset := Shortint(Lo(c1));

  // 2 bytes for offset of last column
  c2 := WordLEToN(AStream.ReadWord);
  ACol2Offset := ShortInt(Lo(c2));

  // Extract info on absolute/relative addresses.
  AFlags := [];
  if (c1 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol);
  if (c1 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow);
  if (c2 and MASK_EXCEL_RELATIVE_COL <> 0) then Include(AFlags, rfRelCol2);
  if (c2 and MASK_EXCEL_RELATIVE_ROW <> 0) then Include(AFlags, rfRelRow2);
end;

procedure TsSpreadBIFF8Reader.ReadRPNSheetIndex(AStream: TStream;
  out ASheet1, ASheet2: Integer);
var
  refIndex: Word;
  ref: TBiff8ExternSheet;
  extbook: TBiff8ExternBook;
begin
  // Index to REF entry in EXTERNSHEET record
  refIndex := WordLEToN(AStream.ReadWord);

  ref := FBiff8ExternSheets[refIndex];
  extBook := FBiff8ExternBooks[ref.ExternBookIndex] as TBiff8ExternBook;

  // Only links to internal sheets supported so far.
  if extBook.Kind <> ebkInternal then
  begin
    ASheet1 := -1;
    ASheet2 := -1;
    exit;
  end;

  ASheet1 := ref.FirstSheetIndex;
  ASheet2 := ref.LastSheetIndex;
end;

procedure TsSpreadBIFF8Reader.ReadRSTRING(AStream: TStream);
var
  j, L: Word;
  ARow, ACol: Cardinal;
  XF: Word;
  wideStrValue: WideString;
  cell: PCell;
  rtfRuns: TBiff8_RichTextFormattingRuns;
  fntIndex: Integer;
  fnt: TsFont;
begin
  { BIFF Record data: Row, Column, XF Index }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Data string: 16-bit length }
  L := WordLEtoN(AStream.ReadWord());

  { Read wide string plus flag, but without processing it }
  wideStrValue := ReadUnformattedWideString(AStream, L);

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);        // "virtual" cell
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);    // "real" cell

  { Save the data string}
  FWorksheet.WriteText(cell, UTF16ToUTF8(wideStrValue));

  { Read rich-text formatting runs }
  L := WordLEToN(AStream.ReadWord);
  SetLength(cell^.RichTextParams, L);
  SetLength(rtfRuns, L);
  AStream.ReadBuffer(rtfRuns[0], L * SizeOf(TBiff8_RichTextFormattingRun));
  for j := 0 to L-1 do
  begin
    // Index of the font. Be aware that the index in the file is not
    // necessarily the same as the index used by the workbook!
    fntIndex := WordLEToN(rtfRuns[j].FontIndex);
    fnt := TsFont(FFontList[fntIndex]);
    fntIndex := FWorkbook.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
    if fntIndex = -1 then
      fntIndex := FWorkbook.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
    cell^.RichTextParams[j].FontIndex := fntIndex;
    // Index of the first character using this font: 0-based in file, 1-based in fps
    cell^.RichTextParams[j].FirstIndex := WordLEToN(rtfRuns[j].FirstIndex) + 1;
    // Hyperlink index - not used by biff
    cell^.RichTextParams[j].HyperlinkIndex := -1;
  end;

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF8Reader.ReadSST(const AStream: TStream);
var
  Items: DWORD;
  StringLength, CurStrLen: WORD;
  LString: String;
  ContinueIndicator: WORD;
  rtParams: TsRichTextParams;
  ms: TMemoryStream;
begin
  //Reads the shared string table, only compatible with BIFF8
  if not Assigned(FSharedStringTable) then begin
    //First time SST creation
    FSharedStringTable := TStringList.Create;

    // Total number of strings in the workbook, not used
    DWordLEtoN(AStream.ReadDWord);

    // Number of following strings
    Items := DWordLEtoN(AStream.ReadDWord);
    Dec(PendingRecordSize, 8);
  end else begin
    //A second record must not happend. Garbage so skip.
    Exit;
  end;

  while Items > 0 do begin
    StringLength := 0;
    StringLength := WordLEtoN(AStream.ReadWord);
    Dec(PendingRecordSize ,2);
    LString := '';

    // This loop takes care of the string being split between the STT and the CONTINUE, or between CONTINUE records
    while PendingRecordSize > 0 do
    begin
      if StringLength > 0 then
        //Read a stream of zero length reads all the stream.
        LString := LString + ReadString(AStream, StringLength, rtParams)
      else
      begin
        //String of 0 chars in length, so just read it empty, reading only the mandatory flags
        AStream.ReadByte; //And discard it.
        Dec(PendingRecordSize);
        //LString:=LString+'';
      end;

      // Check if the record finished and we need a CONTINUE record to go on
      if (PendingRecordSize <= 0) and (Items > 1) then
      begin
        //A Continue will happend, read the
        //tag and continue linking...
        ContinueIndicator := WordLEtoN(AStream.ReadWord);
        if ContinueIndicator <> INT_EXCEL_ID_CONTINUE then begin
          raise EFPSpreadsheetReader.Create('[TsSpreadBIFF8Reader.ReadSST] Expected CONTINUE record not found.');
        end;
        PendingRecordSize := WordLEtoN(AStream.ReadWord);
        CurStrLen := Length(UTF8ToUTF16(LString));
        if StringLength < CurStrLen then
          raise EFPSpreadsheetReader.Create('[TsSpreadBIFF8Reader.ReadSST] StringLength<CurStrLen');
        Dec(StringLength, CurStrLen); //Dec the used chars
        if StringLength = 0 then break;
      end else begin
        break;
      end;
    end;

    if Length(rtParams) = 0 then
      FSharedStringTable.Add(LString)
    else
    begin
      ms := TMemoryStream.Create;
      ms.WriteWord(Length(rtParams));
      ms.WriteBuffer(rtParams[0], SizeOf(TsRichTextParam)*Length(rtParams));
      ms.Position := 0;
      FSharedStringTable.AddObject(LString, ms);
    end;

    {$ifdef FPSPREADDEBUG}
    WriteLn('Adding shared string: ' + LString);
    {$endif}

    dec(Items);
  end;
end;

procedure TsSpreadBIFF8Reader.ReadLabelSST(const AStream: TStream);
var
  ACol,ARow: Cardinal;
  XF: WORD;
  SSTIndex: DWORD;
  rec: TBIFF8_LabelSSTRecord;
  cell: PCell;
  ms: TMemoryStream;
  i, n: Integer;
  rtParams: TsRichTextParams;
  fnt: TsFont;
  fntIndex: Integer;
begin
  rec.Row := 0;  // to silence the compiler...

  { Read entire record, starting at Row }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF8_LabelSSTRecord) - SizeOf(TsBiffHeader));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);
  SSTIndex := DWordLEToN(rec.SSTIndex);

  if SizeInt(SSTIndex) >= FSharedStringTable.Count then begin
    raise EFPSpreadsheetReader.CreateFmt(rsIndexInSSTOutOfRange, [
      Integer(SSTIndex), FSharedStringTable.Count-1
    ]);
  end;

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := FWorksheet.AddCell(ARow, ACol);

  FWorksheet.WriteText(cell, FSharedStringTable.Strings[SSTIndex]);

  { Add attributes }
  ApplyCellFormatting(cell, XF);

  { Add rich text formatting }
  ms := TMemoryStream(FSharedStringTable.Objects[SSTIndex]);
  if ms <> nil then begin
    ms.Position := 0;
    n := WordLEToN(ms.ReadWord);
    SetLength(rtParams, n);
    ms.ReadBuffer(rtParams[0], n*SizeOf(TsRichTextParam));
    SetLength(cell^.RichTextParams, n);
    for i:=0 to n-1 do
    begin
      cell^.RichTextParams[i].FirstIndex := rtParams[i].FirstIndex;
      fntIndex := rtParams[i].FontIndex;
      fnt := TsFont(FFontList[fntIndex]);
      fntIndex := FWorkbook.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
      if fntIndex = -1 then
        fntIndex := FWorkbook.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
      cell^.RichTextParams[i].FontIndex := fntIndex;
      cell^.RichTextParams[i].HyperlinkIndex := -1;
    end;
  end;

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, ARow, ACol, cell);
end;

{ Helper function for reading a string with 8-bit length. }
function TsSpreadBIFF8Reader.ReadString_8bitLen(AStream: TStream): String;
const
  HAS_8BITLEN = true;
var
  wideStr: widestring;
begin
  wideStr := ReadWideString(AStream, HAS_8BITLEN);
//  Result := UTF8Encode(wideStr);   // wp: this leads to string encoding error with fpc 3.0 (no UTF8RTL)
  Result := UTF16ToUTF8(wideStr);
end;

procedure TsSpreadBIFF8Reader.ReadStringRecord(AStream: TStream);
var
  wideStr: WideString;
begin
  wideStr := ReadWideString(AStream, false);
  if (FIncompleteCell <> nil) and (wideStr <> '') then begin
    FIncompleteCell^.UTF8StringValue := UTF8Encode(wideStr);
    FIncompleteCell^.ContentType := cctUTF8String;
    if FIsVirtualMode then
      Workbook.OnReadCellData(Workbook, FIncompleteCell^.Row, FIncompleteCell^.Col, FIncompleteCell);
  end;
  FIncompleteCell := nil;
end;

{ Reads a TXO record (TEXT OBJECT). Needed to retrieve cell comments.
  We only extract the length of the comment text (in characters). The text itself
  is contained in the following CONTINUE record. }
procedure TsSpreadBIFF8Reader.ReadTXO(const AStream: TStream);
var
  rec: TBIFF8TXORecord;
begin
  rec.OptionFlags := 0;  // to silence the compiler
  AStream.ReadBuffer(rec.OptionFlags, Sizeof(Rec) - 2*SizeOf(Word));
  FCommentLen := WordLEToN(rec.TextLen);
end;

procedure TsSpreadBIFF8Reader.ReadXF(const AStream: TStream);
var
  rec: TBIFF8_XFRecord;
  fmt: TsCellFormat;
  b: Byte;
  dw: DWord;
  fill: Integer;
  fs: TsFillStyle;
  nfs: String;
  nfParams: TsNumFormatParams;
  iclr: Integer;
begin
  InitFormatRecord(fmt);
  fmt.ID := FCellFormatList.Count;

  rec.FontIndex := 0;  // to silence the compiler...
  // Read entire xf record into a buffer
  AStream.ReadBuffer(rec.FontIndex, SizeOf(rec) - 2*SizeOf(word));

  // Font index
  fmt.FontIndex := FixFontIndex(WordLEToN(rec.FontIndex));
  if fmt.FontIndex > 1 then
    Include(fmt.UsedFormattingFields, uffFont);

  // Number format index
  if rec.NumFormatIndex <> 0 then begin
    nfs := NumFormatList[rec.NumFormatIndex];
    // "General" (NumFormatIndex = 0) not stored in workbook's NumFormatList
    if (rec.NumFormatIndex > 0) and not SameText(nfs, 'General') then
    begin
      fmt.NumberFormatIndex := Workbook.AddNumberFormat(nfs);
      nfParams := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
      if nfParams <> nil then
      begin
        fmt.NumberFormat := nfParams.NumFormat;
        fmt.NumberFormatStr := nfs;
        Include(fmt.UsedFormattingFields, uffNumberFormat);
      end;
    end;
  end;

  // Horizontal text alignment
  b := rec.Align_TextBreak AND MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
  begin
    fmt.HorAlignment := TsHorAlignment(b);
    if fmt.HorAlignment <> haDefault then
      Include(fmt.UsedFormattingFields, uffHorAlign);
  end;

  // Vertical text alignment
  b := (rec.Align_TextBreak AND MASK_XF_VERT_ALIGN) shr 4;
  if (b + 1 <= ord(high(TsVertAlignment))) then
  begin
    fmt.VertAlignment := TsVertAlignment(b + 1);      // + 1 due to vaDefault
    // Unfortunately BIFF does not provide a "default" vertical alignment code.
    // Without the following correction "non-formatted" cells would always have
    // the uffVertAlign FormattingField set which contradicts the statement of
    // not being formatted.
    if fmt.VertAlignment = vaBottom then
      fmt.VertAlignment := vaDefault;
    if fmt.VertAlignment <> vaDefault then
      Include(fmt.UsedFormattingFields, uffVertAlign);
  end;

  // Word wrap
  if (rec.Align_TextBreak and MASK_XF_TEXTWRAP) <> 0 then
    Include(fmt.UsedFormattingFields, uffWordwrap);

  // BiDi mode
  b := (rec.Indent_Shrink_TextDir and MASK_XF_BIDI) shr 6;
  if b in [0..2] then fmt.BiDiMode := TsBiDiMode(b);
  if b > 0 then Include(fmt.UsedFormattingFields, uffBiDi);

  // TextRotation
  case rec.TextRotation of
    XF_ROTATION_HORIZONTAL : fmt.TextRotation := trHorizontal;
    XF_ROTATION_90DEG_CCW  : fmt.TextRotation := rt90DegreeCounterClockwiseRotation;
    XF_ROTATION_90DEG_CW   : fmt.TextRotation := rt90DegreeClockwiseRotation;
    XF_ROTATION_STACKED    : fmt.TextRotation := rtStacked;
  end;
  if fmt.TextRotation <> trHorizontal then
    Include(fmt.UsedFormattingFields, uffTextRotation);

  // Cell borders
  rec.Border_BkGr1 := DWordLEToN(rec.Border_BkGr1);
  rec.Border_BkGr2 := DWordLEToN(rec.Border_BkGr2);

  // the 4 masked bits encode the line style of the border line. 0 = no line
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_LEFT;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbWest);
    fmt.BorderStyles[cbWest].LineStyle := TsLineStyle(dw-1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_RIGHT;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbEast);
    fmt.BorderStyles[cbEast].LineStyle := TsLineStyle((dw shr 4)-1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_TOP;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbNorth);
    fmt.BorderStyles[cbNorth].LineStyle := TsLineStyle((dw shr 8)-1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr1 and MASK_XF_BORDER_BOTTOM;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbSouth);
    fmt.BorderStyles[cbSouth].LineStyle := TsLineStyle((dw shr 12)-1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border_BkGr2 and MASK_XF_BORDER_DIAGONAL;
  if dw <> 0 then
  begin
    fmt.BorderStyles[cbDiagUp].LineStyle := TsLineStyle((dw shr 21)-1);
    fmt.BorderStyles[cbDiagDown].LineStyle := fmt.BorderStyles[cbDiagUp].LineStyle;
    if rec.Border_BkGr1 and MASK_XF_BORDER_SHOW_DIAGONAL_UP <> 0 then
      Include(fmt.Border, cbDiagUp);
    if rec.Border_BkGr1 and MASK_XF_BORDER_SHOW_DIAGONAL_DOWN <> 0 then
      Include(fmt.Border, cbDiagDown);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;

  // Border line colors
  // NOTE: It is possible that the palette is not yet known at this moment.
  // Therefore we store the palette index encoded into the colorx.
  // They will be converted to rgb in "FixColors".
  iclr := (rec.Border_BkGr1 and MASK_XF_BORDER_LEFT_COLOR) shr 16;
  fmt.BorderStyles[cbWest].Color := IfThen(iclr >= 64, scBlack, SetAsPaletteIndex(iclr));
  iclr := (rec.Border_BkGr1 and MASK_XF_BORDER_RIGHT_COLOR) shr 23;
  fmt.BorderStyles[cbEast].Color := IfThen(iclr >= 64, scBlack, SetAsPaletteIndex(iclr));
  iclr := (rec.Border_BkGr2 and MASK_XF_BORDER_TOP_COLOR);
  fmt.BorderStyles[cbNorth].Color := IfThen(iclr >= 64, scBlack, SetAsPaletteIndex(iclr));
  iclr := (rec.Border_BkGr2 and MASK_XF_BORDER_BOTTOM_COLOR) shr 7;
  fmt.BorderStyles[cbSouth].Color := IfThen(iclr >= 64, scBlack, SetAsPaletteIndex(iclr));
  iclr := (rec.Border_BkGr2 and MASK_XF_BORDER_DIAGONAL_COLOR) shr 14;
  fmt.BorderStyles[cbDiagUp].Color := IfThen(iclr >= 64, scBlack, SetAsPaletteIndex(iclr));
  fmt.BorderStyles[cbDiagDown].Color := fmt.BorderStyles[cbDiagUp].Color;

  // Background fill pattern and color
  fill := (rec.Border_BkGr2 and MASK_XF_BACKGROUND_PATTERN) shr 26;
  if fill <> MASK_XF_FILL_PATT_EMPTY then
  begin
    rec.BkGr3 := DWordLEToN(rec.BkGr3);
    for fs in TsFillStyle do
      if fill = MASK_XF_FILL_PATT[fs] then
      begin
        // Pattern color
        iclr := rec.BkGr3 and $007F;
        fmt.Background.FgColor := IfThen(iclr = SYS_DEFAULT_FOREGROUND_COLOR,
          scBlack, SetAsPaletteIndex(iclr));

        // Background color
        iclr := (rec.BkGr3 and $3F80) shr 7;
        fmt.Background.BgColor := IfThen(iclr = SYS_DEFAULT_BACKGROUND_COLOR,
          scTransparent, SetAsPaletteIndex(iclr));

        // Fill style
        fmt.Background.Style := fs;
        Include(fmt.UsedFormattingFields, uffBackground);
        break;
      end;
  end;

  // Protection
  case WordLEToN(rec.XFType_Prot_ParentXF) and MASK_XF_TYPE_PROTECTION of
    0:
      fmt.Protection := [];
    MASK_XF_TYPE_PROT_LOCKED:
      fmt.Protection := [cpLockCell];
    MASK_XF_TYPE_PROT_FORMULA_HIDDEN:
      fmt.Protection := [cpHideFormulas];
    MASK_XF_TYPE_PROT_LOCKED + MASK_XF_TYPE_PROT_FORMULA_HIDDEN:
      fmt.Protection := [cpLockCell, cpHideFormulas];
  end;
  if fmt.Protection <> DEFAULT_CELL_PROTECTION then
    Include(fmt.UsedFormattingFields, uffProtection);

  // Add the XF to the internal cell format list
  FCellFormatList.Add(fmt);
end;

{ Reads a DEFINEDNAME record. Currently only extract print ranges and titles. }
procedure TsSpreadBIFF8Reader.ReadDEFINEDNAME(const AStream: TStream);
var
  options: Word;
  len: byte;
  formulaSize: Word;
  widestr: WideString;
  defName: String;
  rpnformula: TsRPNFormula;
  rtf: TsRichTextParams;
  validOnSheet: Integer;
begin
  // Options
  options := WordLEToN(AStream.ReadWord);
  if options and $0020 = 0 then   // only support built-in names at the moment!
    exit;

  // Keyboard shortcut  --> ignore
  AStream.ReadByte;

  // Length of name (character count)
  len := AStream.ReadByte;

  // Size of formula data
  formulasize := WordLEToN(AStream.ReadWord);

  // not used
  AStream.ReadWord;

  // Sheet index (1-based) on which the name is valid (0 = global)
  validOnSheet := SmallInt(WordLEToN(AStream.ReadWord)) - 1;  // now 0-based!

  // Length of Menu text (ignore)
  AStream.ReadByte;

  // Length of description text(ignore)
  AStream.ReadByte;

  // Length of help topic text (ignore)
  AStream.ReadByte;

  // Length of status bar text (ignore)
  AStream.ReadByte;

  // Name
  wideStr := ReadWideString(AStream, len, rtf);
  defName := UTF8Encode(widestr);

  // Formula
  if not ReadRPNTokenArray(AStream, formulaSize, rpnFormula) then
    exit;
  // Store defined name in internal list
  FDefinedNames.Add(TsBIFFDefinedName.Create(defName, rpnFormula, validOnSheet));

  // Skip rest...
end;

procedure TsSpreadBIFF8Reader.ReadEXTERNBOOK(const AStream: TStream);
var
  i, n: Integer;
  url: widestring;
  sheetnames: widestring;
  externbook: TBiff8Externbook;
  p: Int64;
  t: array[0..1] of byte;
begin
  if FBiff8ExternBooks = nil then
    FBiff8ExternBooks := TFPObjectList.Create(true);

  externBook := TBiff8ExternBook.Create;

  // Count of sheets in book
  n := WordLEToN(AStream.ReadWord);

  // Determine type of book
  p := AStream.Position;
  AStream.ReadBuffer(t[0], 2);
  if (t[0] = 1) and (t[1] = 4) then
    externbook.Kind := ebkInternal
  else
  if (t[0] = 1) and (t[1] = $3A) then
    externbook.Kind := ebkAddInFunc
  else
  if n = 0 then
    externbook.Kind := ebkDDE_OLE
  else
    externbook.Kind := ebkExternal;

  if (externbook.Kind = ebkExternal) then
  begin
    AStream.Position := p;

    // Encoded URL without sheet name (Unicode string, 16bit string length)
    url := ReadWideString(AStream, false);
    externbook.DocumentURL := UTF8Encode(url);

    if n = 0 then
      sheetnames := ''
    else begin
      // Sheet names (Unicode strings with 16bit string length)
      sheetnames := UTF8Encode(ReadWideString(AStream, false));
      for i := 2 to n do
        sheetnames := sheetnames + #1 + UTF8Encode(ReadWideString(AStream, false));
    end;
    externbook.SheetNames := sheetNames;
  end;

  FBiff8ExternBooks.Add(externbook);
end;

{ Reads an EXTERNSHEET record. Needed for 3d-references, named cells and
  print ranges. }
procedure TsSpreadBIFF8Reader.ReadEXTERNSHEET(const AStream: TStream);
var
  numItems: Word;
  i: Integer;
begin
  numItems := WordLEToN(AStream.ReadWord);
  SetLength(FBiff8ExternSheets, numItems);

  for  i := 0 to numItems-1 do begin
    AStream.ReadBuffer(FBiff8ExternSheets[i], Sizeof(FBiff8ExternSheets[i]));
    with FBiff8ExternSheets[i] do
    begin
      ExternBookIndex := WordLEToN(ExternBookIndex);
      FirstSheetIndex := WordLEToN(FirstSheetIndex);
      LastSheetIndex := WordLEToN(LastSheetIndex);
    end;
  end;
end;


{ Reads a FONT record. The retrieved font is stored in the workbook's FontList. }
procedure TsSpreadBIFF8Reader.ReadFONT(const AStream: TStream);
var
  {%H-}lCodePage: Word;
  lHeight: Word;
  lOptions: Word;
  lColor: Word;
  lWeight: Word;
  Len: Byte;
  font: TsFont;
  rtParams: TsRichTextParams;
  isDefaultFont: Boolean;
begin
  font := TsFont.Create;

  { Height of the font in twips = 1/20 of a point }
  lHeight := WordLEToN(AStream.ReadWord);
  font.Size := lHeight / 20;

  { Option flags }
  lOptions := WordLEToN(AStream.ReadWord);
  font.Style := [];
  if lOptions and $0001 <> 0 then Include(font.Style, fssBold);
  if lOptions and $0002 <> 0 then Include(font.Style, fssItalic);
  if lOptions and $0004 <> 0 then Include(font.Style, fssUnderline);
  if lOptions and $0008 <> 0 then Include(font.Style, fssStrikeout);

  { Color index }
  // The problem is that the palette is loaded after the font list; therefore
  // we do not know the rgb color of the font here. We store the palette index
  // ("SetAsPaletteIndex") and replace it by the rgb color after reading of the
  // palette and after reading the workbook globals records. As an indicator
  // that the font does not yet contain an rgb color a control bit is set in
  // the high-byte of the TsColor.
  lColor := WordLEToN(AStream.ReadWord);
  if lColor < 8 then
    // Use built-in colors directly otherwise the Workbook's FindFont would not find the font in ReadXF
    font.Color := FPalette[lColor]
  else
  if lColor = SYS_DEFAULT_WINDOW_TEXT_COLOR then
    font.Color := scBlack
  else
    font.Color := SetAsPaletteIndex(lColor);

  { Font weight }
  lWeight := WordLEToN(AStream.ReadWord);
  if lWeight = 700 then Include(font.Style, fssBold);

  { Escapement type }
  font.Position := TsFontPosition(WordLEToN(AStream.ReadWord));

  { Underline type }
  if AStream.ReadByte > 0 then Include(font.Style, fssUnderline);

  { Font family }
  AStream.ReadByte();

  { Character set }
  lCodepage := AStream.ReadByte();
  {$ifdef FPSPREADDEBUG}
  WriteLn('Reading Font Codepage='+IntToStr(lCodepage));
  {$endif}

  { Not used }
  AStream.ReadByte();

  { Font name: Unicodestring, char count in 1 byte }
  Len := AStream.ReadByte();
  font.FontName := ReadString(AStream, Len, rtParams);  // rtParams is not used here.

  isDefaultFont := FFontList.Count = 0;

  { Add font to internal font list; will be transferred to workbook later because
    the font index in the internal list (= index in file) is not the same as the
    index the font will have in the workbook's fontlist! }
  FFontList.Add(font);

  { Excel does not have zero-based font #4! }
  if FFontList.Count = 4 then FFontList.Add(nil);

  if isDefaultFont then
    Workbook.SetDefaultFont(font.FontName, font.Size);
end;

{@@ ----------------------------------------------------------------------------
  Reads the (number) FORMAT record for formatting numerical data and stores the
  format strings in an internal stringlist. The strings are put at the index
  specified by the FORMAT record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.ReadFORMAT(AStream: TStream);
var
  fmtString: String;
  fmtIndex: Integer;
begin
  // Record FORMAT, BIFF 8 (5.49):
  // Offset Size Contents
  // 0      2     Format index used in other records
  // 2      var   Number format string (Unicode string, 16-bit string length)
  // From BIFF5 on: indexes 0..163 are built in
  fmtIndex := WordLEtoN(AStream.ReadWord);
  if fmtIndex = 0 then  // "General" already in list
    exit;

  // 2 var. Number format string (Unicode string, 16-bit string length, ➜2.5.3)
//  fmtString := UTF8Encode(ReadWideString(AStream, False));
  fmtString := UTF16ToUTF8(ReadWideString(AStream, False));

  // Add to the list at the specified index. If necessary insert empty strings
  while NumFormatList.Count <= fmtIndex do NumFormatList.Add('');
  NumFormatList[fmtIndex] := fmtString;
end;

{@@ ----------------------------------------------------------------------------
  Reads the header/footer to be used for printing.
  Overriden for BIFF8 because of wide strings
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.ReadHeaderFooter(AStream: TStream;
  AIsHeader: Boolean);
var
  s: widestring;
  len: word;
  rtParams: TsRichTextParams;
begin
  if RecordSize = 0 then
    exit;

  len := WordLEToN(AStream.ReadWord);
  s := ReadWideString(AStream, len, rtParams);
  if AIsHeader then
    FWorksheet.PageLayout.Headers[1] := UTF8Encode(s)
  else
    FWOrksheet.PageLayout.Footers[1] := UTF8Encode(s);

  { Options poDifferentFirst and poDifferentOddEvent are not used, BIFF supports
    only common headers/footers }
end;

{@@ ----------------------------------------------------------------------------
  Reads a HYPERLINK record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.ReadHyperlink(const AStream: TStream);
var
  row, col, row1, col1, row2, col2: word;
  guid: TGUID;
  flags: DWord;
  widestr: widestring;
  len: DWord;
  link: String;
  linkDos: String;
  mark: String;
  dirUpCount: Word;
  ansistr: ansistring;
  size: DWord;
begin
  { Row and column index range of cells using the hyperlink }
  row1 := WordLEToN(AStream.ReadWord);
  row2 := WordLEToN(AStream.ReadWord);
  col1 := WordLEToN(AStream.ReadWord);
  col2 := WordLEToN(AStream.ReadWord);

  { GUID of standard link }
  AStream.ReadBuffer(guid{%H-}, SizeOf(guid));

  { unknown DWord }
  AStream.ReadDWord;

  { Flags }
  flags := DWordLEToN(AStream.ReadDWord);

  { Description }
  if flags and MASK_HLINK_DESCRIPTION = MASK_HLINK_DESCRIPTION then
  begin
    // not used because there is always a "normal" cell to which the hyperlink is associated.
    // character count of description incl trailing zero
    len := DWordLEToN(AStream.ReadDWord);
    // Character array (16-bit characters, with trailing zero word)
    SetLength(wideStr, len);
    AStream.ReadBuffer(wideStr[1], len*SizeOf(wideChar));
  end;

  { Target frame: external link (URI or local file) }
  link := '';
  if flags and MASK_HLINK_LINK <> 0 then
  begin
    AStream.ReadBuffer(guid, SizeOf(guid));

    // Check for URL
    if GuidToString(guid) = '{79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}' then
    begin
      // Size of character array incl trailing zero
      size := DWordLEToN(AStream.ReadDWord);
      // Character array of URL (16-bit-characters, with trailing zero word)
      // See 3 lines below: This buffer is too large!
      len := size div 2 - 1;
      SetLength(wideStr, len);
      AStream.ReadBuffer(wideStr[1], size);
      // The buffer can be larger than the space occupied by the wideStr.
      // --> Find true string length and convert wide string to utf-8.

//      len := StrLen(PWideChar(widestr));  // wp: working fine except for Laz1.0
      len := Length(widestr);               // Is this ok?

      SetLength(widestr, len);
      link := UTF8Encode(widestr);
    end else
    // Check for local file
    if GuidToString(guid) = '{00000303-0000-0000-C000-000000000046}' then
    begin
      dirUpCount := WordLEToN(AStream.ReadWord);
      // Character count of the shortened file path and name, incl trailing zero byte
      len := DWordLEToN(AStream.ReadDWord);
      // Character array of the shortened file path and name in 8.3-DOS-format.
      // This field can be filled with a long file name too.
      // No unicode string header, always 8-bit characters, zero-terminated.
      SetLength(ansiStr, len);
      AStream.ReadBuffer(ansiStr[1], len*SizeOf(ansiChar));
      SetLength(ansistr, len-1);  // Remove trailing zero
      while dirUpCount > 0 do
      begin
        ansistr := '..' + PathDelim + ansistr;
        dec(dirUpCount);
      end;
      linkDos := AnsiToUTF8(ansiStr);
      // 6 unknown DWord values
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      AStream.ReadDWord;
      // Size of the following file link field including string length field
      // and additional data field
      size := DWordLEToN(AStream.ReadDWord);
      if size > 0 then
      begin
        // Size of the extended file path and name.
        size := DWordLEToN(AStream.ReadDWord);
        len := size div 2;
        // Unknown
        AStream.ReadWord;
        // Character array of the extended file path and name
        // no Unicode string header, always 16-bit characters, not zero-terminated
        SetLength(wideStr, len);
        AStream.ReadBuffer(wideStr[1], size);
        link := UTF8Encode(wideStr);
      end else
        link := linkDos;

      // An absolute path must be a fully qualified URI to be compatible with fps
      if flags and MASK_HLINK_ABSOLUTE <> 0 then
        link := FilenameToURI(link);
    end;
  end;

  { Text mark }
  if flags and MASK_HLINK_TEXTMARK = MASK_HLINK_TEXTMARK then
  begin
    // Character count of the text mark, including trailing zero word
    len := DWordLEToN(AStream.ReadDWord);
    // Character array of the text mark without "#" sign
    // no Unicode string header, always 16-bit characters, zero-terminated
    SetLength(wideStr, len);
    AStream.ReadBuffer(wideStr[1], len*SizeOf(wideChar));
    SetLength(wideStr, len-1);  // Remove trailing zero word
    mark := UTF8Encode(wideStr);
  end;

  // Add bookmark to hyperlink target
  if (link <> '') and (mark <> '') then
    link := link + '#' + mark
  else
  if (link = '') then
    link := '#' + mark;

  // Add hyperlink(s) to worksheet
  for row := row1 to row2 do
    for col := col1 to col2 do
      FWorksheet.WriteHyperlink(row, col, link);
end;


{@@ ----------------------------------------------------------------------------
  Reads a HYPERLINK TOOLTIP record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Reader.ReadHyperlinkToolTip(const AStream: TStream);
var
  txt: String;
  widestr: widestring;
  //row, col,
  row1, col1, row2, col2: Word;
  hyperlink: PsHyperlink;
  numbytes: Integer;
begin
  { Record type; this matches the BIFF record type }
  AStream.ReadWord;

  { Row and column index range of cells using the hyperlink tooltip }
  row1 := WordLEToN(AStream.ReadWord);
  row2 := WordLEToN(AStream.ReadWord);
  col1 := WordLEToN(AStream.ReadWord);
  col2 := WordLEToN(AStream.ReadWord);

  { Hyperlink tooltip, a null-terminated unicode string }
  numbytes := RecordSize - 5*SizeOf(word);
  SetLength(wideStr, numbytes div 2);
  AStream.ReadBuffer(wideStr[1], numbytes);
  SetLength(wideStr, Length(wideStr)-1);  // Remove trailing zero word
  txt := UTF8Encode(wideStr);

  { Add tooltip to hyperlinks }
  for hyperlink in FWorksheet.Hyperlinks.GetRangeEnumerator(row1, col1, row2, col2) do
    hyperlink^.ToolTip := txt;
end;


{------------------------------------------------------------------------------}
{                          TsSpreadBIFF8Writer                                 }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the Excel 8 writer
-------------------------------------------------------------------------------}
constructor TsSpreadBIFF8Writer.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  InitBiff8Limitations(FLimitations);
  FDateMode := Excel8Settings.DateMode;
  PopulateSharedStringTable(AWorkbook);
end;

destructor TsSpreadBIFF8Writer.Destroy;
begin
  FSharedStringTable.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Writes the header of a CONTINUE record. Since the size of the record is not
  known at this moment the position of the BIFF size field is returned in
  ASizePos. Later the size field must be corrected by calling FixRecordSize
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.BeginCONTINUERecord(AStream: TStream;
  out ASizePos: Int64);
begin
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_CONTINUE));
  ASizePos := AStream.Position;
  AStream.WriteWord(0);
end;

{@@ ----------------------------------------------------------------------------
  Sometimes the size of records is not known when writing them (see
  BeginCONTINUERecord). This method rewinds the stream to the position where
  the record size is expected (ASizePos) and writes the record size (ASize).
  Then the stream returns to its original position.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.FixRecordSize(AStream: TStream; ASizePos: Int64;
  ASize: Word);
var
  p: Int64;
begin
  p := AStream.Position;
  AStream.Position := ASizePos;
  AStream.WriteWord(WordToLE(ASize));
  AStream.Position := p;
end;

function TsSpreadBIFF8Writer.GetPrintOptions: Word;
Begin
  Result := inherited GetPrintOptions;
{ The following flags are valid for BIFF8 only:
  Bit 9: 0 = Print notes as displayed; 1 = Print notes at end of sheet
  Bit 11-10:  00 = Print errors as displayed; 1 = Do not print errors
              2 = Print errors as “--”; 3 = Print errors as “#N/A” }
  if poCommentsAtEnd in FWorksheet.PageLayout.Options then
    Result := Result or $0200;
end;

{@@ ----------------------------------------------------------------------------
  If the specified string and richtextparams are stored in the SharedStringTable
  then the index is returned, otherwise -1
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.IndexOfSharedString(const AText: String;
  const ARichTextParams: TsRichTextParams): Integer;
var
  s: String;
  obj: TObject;
begin
  if FSharedStringTable <> nil then
    for Result := 0 to FSharedStringTable.Count-1 do begin
      s := FSharedStringTable.Strings[Result];
      obj := FSharedStringTable.Objects[Result];
//      if (s = AText) and (TsRichTextParams(obj) = ARichTextParams)
      if (s = AText) and SameRichTextParams(TsRichTextParams(obj), ARichTextParams)
        then exit;
    end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel BIFF8 record structure to a stream

  Be careful as this method doesn't write the OLE part of the document,
  just the BIFF records
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.InternalWriteToStream(AStream: TStream);
const
  isBIFF8 = true;
var
  currentPos: Int64;
  sheetPos: array of Int64;
  i: Integer;
  pane: Byte;
begin
  { Write workbook globals }
  WriteBOF(AStream, INT_BOF_WORKBOOK_GLOBALS);
  WriteCodePage(AStream, 'ucs2le'); // = utf-16
  WriteWindowProtect(AStream, bpLockWindows in Workbook.Protection);
  WritePROTECT(AStream, bpLockStructure in Workbook.Protection);
  WritePASSWORD(AStream, Workbook.CryptoInfo);
  WriteWINDOW1(AStream);
  WriteFonts(AStream);
  WriteNumFormats(AStream);
  WritePalette(AStream);
  WriteXFRecords(AStream);
  WriteStyle(AStream);

  // A BOUNDSHEET for each worksheet
  SetLength(sheetPos, Workbook.GetWorksheetCount);
  for i := 0 to Workbook.GetWorksheetCount - 1 do
    sheetPos[i] := WriteBoundsheet(AStream, Workbook.GetWorksheetByIndex(i));

  WriteEXTERNBOOK(AStream);
  WriteEXTERNSHEET(AStream);
  WriteDefinedNames(AStream);
  WriteSST(AStream);

  WriteEOF(AStream);

  { Write each worksheet }
  for i := 0 to Workbook.GetWorksheetCount - 1 do
  begin
    FWorksheet := Workbook.GetWorksheetByIndex(i);

    { First goes back and writes the position of the BOF of the
      sheet on the respective BOUNDSHEET record }
    currentPos := AStream.Position;
    AStream.Position := sheetPos[i];
    AStream.WriteDWord(DWordToLE(DWORD(CurrentPos)));
    AStream.Position := currentPos;

    WriteBOF(AStream, INT_BOF_SHEET);
      WriteIndex(AStream);
      WritePrintHeaders(AStream);
      WritePrintGridLines(AStream);
      WriteDefaultRowHeight(AStream, FWorksheet);
      WriteSheetPR(AStream);

      // Page setting block
      WriteHeaderFooter(AStream, true);
      WriteHeaderFooter(AStream, false);
      WriteHCenter(AStream);
      WriteVCenter(AStream);
      WriteMargin(AStream, 0);  // 0 = left margin
      WriteMargin(AStream, 1);  // 1 = right margin
      WriteMargin(AStream, 2);  // 2 = top margin
      WriteMargin(AStream, 3);  // 3 = bottom margin
      WritePageSetup(AStream);

      // Protection
      if FWorksheet.IsProtected then begin
        WritePROTECT(AStream, true);
//        WriteScenarioProtect(AStream);
        WriteObjectProtect(AStream, FWorksheet);
        WritePASSWORD(AStream, FWorksheet.CryptoInfo);
      end;

      WriteDefaultColWidth(AStream, FWorksheet);
      WriteColInfos(AStream, FWorksheet);
      WriteDimensions(AStream, FWorksheet);

      if (boVirtualMode in Workbook.Options) then
        WriteVirtualCells(AStream, FWorksheet)
      else begin
        WriteRows(AStream, FWorksheet);
        WriteCellsToStream(AStream, FWorksheet.Cells);
        WriteComments(AStream, FWorksheet);
      end;

      // View settings block
      WriteWindow2(AStream, FWorksheet);
      WriteSCLRecord(AStream, FWorksheet);
      WritePane(AStream, FWorksheet, isBIFF8, pane);
      WriteSelection(AStream, FWorksheet, pane);
      WriteHyperlinks(AStream, FWorksheet);

      WriteMergedCells(AStream, FWorksheet);

    WriteEOF(AStream);
  end;

  { Cleanup }
  SetLength(sheetPos, 0);
end;

{@@ ----------------------------------------------------------------------------
  Populates the palette of the writer with the colors used by the workbook.
  BIFF8 begins with the 8 default colors which are duplicated. Then the user
  colors follow up to a max of total 64 entries.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.PopulatePalette(AWorkbook: TsWorkbook);
var
  i: Integer;
begin
  FPalette.Clear;
  FPalette.AddBuiltinColors(true);
  FPalette.CollectFromWorkbook(AWorkbook);

  { Fill up Excel colors of the standard palette to avoid empty color
    place holders in Excel's colordialog. }
  i := 16;
  while (i <= High(PALETTE_BIFF8)) and (FPalette.Count < 64) do begin
    FPalette.AddColor(PALETTE_BIFF8[i]);
    // The BIFF8 palette contains duplicate colors -> don't use AddUniqueColor
    inc(i);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Collects all strings of the workbook in the SharedStringTable
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.PopulateSharedStringTable(AWorkbook: TsWorkbook);
var
  i, idx: Integer;
  cell: PCell;
  sheet: TsWorksheet;
begin
  FNumStrings := 0;
  FSharedStringTable := TStringList.Create;

  for i:=0 to AWorkbook.GetWorksheetCount-1 do
  begin
    sheet := AWorkbook.GetWorksheetByIndex(i);
    for cell in sheet.Cells do begin
      if (cell^.ContentType <> cctUTF8String) then
        Continue;
      if (cell^.UTF8StringValue = '') then
        Continue;
      idx := IndexOfSharedString(cell^.UTF8StringValue, cell^.RichTextParams);
      inc(FNumStrings);
      if idx > -1 then
        Continue;
      FSharedStringTable.AddObject(cell^.UTF8StringValue, TObject(cell^.RichTextParams));
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel BIFF8 file to the disc

  The BIFF 8 writer overrides this method because BIFF 8 is written
  as an OLE document, and our current OLE document writing method involves:

         1 - Writing the BIFF data to a memory stream
         2 - Write the memory stream data to disk using COM functions
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean; AParams: TsStreamParams = []);
var
  Stream: TStream;
  OutputStorage: TOLEStorage;
  OLEDocument: TOLEDocument;
begin
  Unused(AParams);
  if (boBufStream in Workbook.Options) then begin
    Stream := TBufStream.Create
  end else
    Stream := TMemoryStream.Create;
  try
    InternalWriteToStream(Stream);
    OutputStorage := TOLEStorage.Create;
    try
      // Only one stream is necessary for any number of worksheets
      OLEDocument.Stream := Stream;
      OutputStorage.WriteOLEFile(AFileName, OLEDocument, AOverwriteExisting, 'Workbook');
    finally
      OutputStorage.Free;
    end;
  finally
    Stream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel BIFF8 record structure to a stream containing the OLE
  envelope of the document.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteToStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  OutputStorage: TOLEStorage;
  OLEDocument: TOLEDocument;
  stream: TStream;
begin
  Unused(AParams);

  if (boBufStream in Workbook.Options) then
    stream := TBufStream.Create else
    stream := TMemoryStream.Create;
  try
    InternalWriteToStream(stream);
    OutputStorage := TOLEStorage.Create;
    try
      // Only one stream is necessary for any number of worksheets
      OLEDocument.Stream := stream;
      OutputStorage.WriteOLEStream(AStream, OLEDocument, 'Workbook');
    finally
      OutputStorage.Free;
    end;
  finally
    stream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 BOF record

  This must be the first record on an Excel 8 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteBOF(AStream: TStream; ADataType: Word);
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_BOF, 16);

  { BIFF version. Should only be used if this BOF is for the workbook globals }
  { OpenOffice rejects to correctly read xls files if this field is
    omitted as docs. says, or even if it is being written to zero value,
    Not tested with Excel, but MSExcel reader opens it as expected }
  AStream.WriteWord(WordToLE(INT_BOF_BIFF8_VER));

  { Data type }
  AStream.WriteWord(WordToLE(ADataType));

  { Build identifier, must not be 0 }
  AStream.WriteWord(WordToLE(INT_BOF_BUILD_ID));

  { Build year, must not be 0 }
  AStream.WriteWord(WordToLE(INT_BOF_BUILD_YEAR));

  { File history flags }
  AStream.WriteDWord(DWordToLE(0));

  { Lowest Excel version that can read all records in this file 5?}
  AStream.WriteDWord(DWordToLE(0)); //?????????
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 BOUNDSHEET record
  Always located in the workbook globals substream.
  One BOUNDSHEET is written for each worksheet.

  @return   The stream position where the absolute stream position
            of the BOF of this sheet should be written (4 bytes size).
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.WriteBoundsheet(AStream: TStream;
  AWorksheet: TsWorksheet): Int64;
var
  len: Byte;
  wideSheetName: WideString;
  sheetState: Byte;
begin
  wideSheetName := UTF8Decode(AWorksheet.Name);
  len := Length(wideSheetName);

  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_BOUNDSHEET, 8 + Len * Sizeof(WideChar));

  { Absolute stream position of the BOF record of the sheet represented
    by this record }
  Result := AStream.Position;
  AStream.WriteDWord(DWordToLE(0));

  { Visibility }
  sheetState := IfThen(soHidden in AWorksheet.Options, 1, 0);
  AStream.WriteByte(sheetState);

  { Sheet type }
  AStream.WriteByte(0);

  { Sheet name: Unicode string char count 1 byte }
  AStream.WriteByte(Len);
  {String flags}
  AStream.WriteByte(1);
  AStream.WriteBuffer(WideStringToLE(wideSheetName)[1], len * SizeOf(WideChar));
end;

{@@ ----------------------------------------------------------------------------
  Inherited method for writing a cell comment immediately after cell content.
  A writing method has been implemented by xlscommon. But in BIFF8, this
  must not do anything because comments are collected in a list and
  written en-bloc. See WriteComments.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteComment(AStream: TStream; ACell: PCell);
begin
  // Nothing to do. Reverts the behavior introduced by xlscommon.
  Unused(AStream, ACell);
end;

{@@ ----------------------------------------------------------------------------
  Writes all comments to the worksheet stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteComments(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  index: Integer;
  comment: PsComment;
begin
  exit;      // Remove after comments can be written correctly
  {$warning TODO: Fix writing of cell comments in BIFF8 (file is readable by OpenOffice, but not by Excel)}

  { At first we have to write all Escher-related records for all comments;
      MSODRAWING - OBJ - MSODRAWING - TXO }
  index := 1;
  for comment in AWorksheet.Comments do
  begin
    if index = 1 then
      WriteMSODrawing1(AStream, FWorksheet.Comments.Count, comment)
    else
      WriteMSODrawing2(AStream, comment, index);
    WriteOBJ(AStream, index);
    WriteMSODrawing3(AStream);
    WriteTXO(AStream, comment);
    inc(index);
  end;

  { The NOTE records for all comments follow subsequently. }
  index := 1;
  for comment in AWorksheet.Comments do
  begin
    WriteNOTE(AStream, comment, index);
    inc(index);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a DEFINEDNAME record.
  Implements only the builtin defined names for print ranges and titles!
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteDefinedName(AStream: TStream;
  AWorksheet: TsWorksheet; const AName: String; AIndexToREF: Word);

  procedure WriteRangeFormula(MemStream: TMemoryStream; ARange: TsCellRange;
    AIndexToRef, ACounter: Word);
  begin
    { Token for tArea3dR }
    MemStream.WriteByte($3B);

    { Index to REF entry in EXTERNSHEET record }
    MemStream.WriteWord(WordToLE(AIndexToREF));

    { First row index }
    MemStream.WriteWord(WordToLE(ARange.Row1));

    { Last row index }
    MemStream.WriteWord(WordToLE(ARange.Row2));

    { First column index }
    MemStream.WriteWord(WordToLE(ARange.Col1));

    { Last column index }
    MemStream.WriteWord(WordToLE(ARange.Col2));

    { Token for list if formula refers to more than 1 range }
    if ACounter > 1 then
      MemStream.WriteByte($10);
  end;

var
  memstream: TMemoryStream;
  rng: TsCellRange;
  j: Integer;

begin
  // Since this is a variable length record we begin by writing the formula
  // to a memory stream
  memstream := TMemoryStream.Create;
  try
    case AName of
      #06: begin  // Print range
             for j := 0 to AWorksheet.PageLayout.NumPrintRanges-1 do
             begin
               rng := AWorksheet.PageLayout.PrintRange[j];
               WriteRangeFormula(memstream, rng, AIndexToRef, j+1);
             end;
           end;
      #07: begin  // Print titles
             j := 1;
             if AWorksheet.PageLayout.HasRepeatedCols then
             begin
               rng.Col1 := AWorksheet.PageLayout.RepeatedCols.FirstIndex;
               rng.Col2 := AWorksheet.PageLayout.RepeatedCols.LastIndex;
               if rng.Col2 = UNASSIGNED_ROW_COL_INDEX then rng.Col2 := rng.Col1;
               rng.Row1 := 0;
               rng.Row2 := 65535;
               WriteRangeFormula(memstream, rng, AIndexToRef, j);
               inc(j);
             end;
             if AWorksheet.PageLayout.HasRepeatedRows then
             begin
               rng.Row1 := AWorksheet.PageLayout.RepeatedRows.FirstIndex;
               rng.Row2 := AWorksheet.PageLayout.RepeatedRows.LastIndex;
               if rng.Row2 = UNASSIGNED_ROW_COL_INDEX then rng.Row2 := rng.Row1;
               rng.Col1 := 0;
               rng.Col2 := 255;
               WriteRangeFormula(memstream, rng, AIndexToRef, j);
             end;
           end;
      else
        raise EFPSpreadsheetWriter.Create('Name not supported');
    end;  // case

    { BIFF record header }
    WriteBIFFHeader(AStream, INT_EXCEL_ID_DEFINEDNAME, 16 + memstream.Size);
    // NOTE: 16 only valid for internal names !!!!

    { Option flags: built-in defined names only }
    AStream.WriteWord(WordToLE($0020));

    { Keyboard shortcut (only for command macro names) }
    AStream.WriteByte(0);

    { Length of name (character count). Always 1 for builtin names }
    AStream.WriteByte(1);

    { Size of formula data }
    AStream.WriteWord(WordToLE(memstream.Size));

    { not used }
    AStream.WriteWord(0);

    { Index to sheet (1-based) }
    AStream.WriteWord(WordToLE(FWorkbook.GetWorksheetIndex(AWorksheet)+1));

    { Length of menu text }
    AStream.WriteByte(0);

    { Length of description text }
    AStream.WriteByte(0);

    { Length of help topic text }
    AStream.WriteByte(0);

    { Length of status bar text }
    AStream.WriteByte(0);

    { Name }
    if (Length(AName) = 1) and (AName[1] < #32) then
      AStream.WriteWord(WordToLE(ord(AName[1]) shl 8))
    else
      raise EFPSpreadsheetWriter.Create('Name not supported.');

    { Formula }
    memstream.Position := 0;
    AStream.CopyFrom(memstream, memstream.Size);

  finally
    memstream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 DIMENSIONS record

  nm = (rl - rf - 1) / 32 + 1 (using integer division)

  Excel, OpenOffice and FPSpreadsheet ignore the dimensions written in this
  record, but some other applications really use them, so they need to be correct.

  See bug 18886: excel5 files are truncated when imported
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteDimensions(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  firstRow, lastRow, firstCol, lastCol: Cardinal;
  rec: TBIFF8_DimensionsRecord;
begin
  { Determine sheet size }
  GetSheetDimensions(AWorksheet, firstRow, lastRow, firstCol, lastCol);

  { Populate BIFF record }
  rec.RecordID := WordToLE(INT_EXCEL_ID_DIMENSIONS);
  rec.RecordSize := WordToLE(14);
  rec.FirstRow := DWordToLE(firstRow);
  rec.LastRowPlus1 := DWordToLE(lastRow+1);
  rec.FirstCol := WordToLE(firstCol);
  rec.LastColPlus1 := WordToLE(lastCol+1);
  rec.NotUsed := 0;

  { Write BIFF record to stream }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 EOF record.
  This must be the last record on an Excel 8 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_EOF, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an EXTERNBOOK record needed for defined names and links.
  NOTE: This writes only the case for "internal references" required for print
  ranges and titles.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteEXTERNBOOK(AStream: TStream);
begin
  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_EXTERNBOOK, 4);

  { Number of sheets in this workbook }
  AStream.WriteWord(WordToLE(FWorkbook.GetWorksheetCount));

  { Relict from BIFF5 }
  AStream.WriteWord(WordToLE($0401));
end;

{@@ ----------------------------------------------------------------------------
  Writes an EXTERNSHEET record needed for defined names and links.
  NOTE: This writes only what is required for print ranges and titles.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteEXTERNSHEET(AStream: TStream);
var
  n, i: Integer;
begin
  { Since sheet range are not supported we simply note every sheet here. }
  n := FWorkbook.GetWorksheetCount;

  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_EXTERNSHEET, 2 + 6*n);

  { Count of following REF structures }
  AStream.WriteWord(WordToLE(n));

  { REF record for each sheet }
  for i := 0 to n-1 do
  begin
    AStream.WriteWord(0);            // Index to EXTERNBOOK record, always 0
    AStream.WriteWord(WordToLE(i));  // Index to first sheet in EXTERNBOOK sheet list
    AStream.WriteWord(WordToLE(i));  // Index to last sheet in EXTERNBOOK sheet list
  end;
end;
(*

  write a record for
  // every sheet
type
  TExternRefRec = record
    FirstIndex, LastIndex: Word;
  end;
const
  BUF_COUNT = 10;
var
  extern: Array of TExternRefRec;
  sheet: TsWorksheet;
  cell: PCell;
  i, j: Integer;
  n: Word;
  writeIt: Boolean;
begin

  n := 0;
  SetLength(extern, BUF_COUNT);

  // Find sheets used in formula references
  for i:=0 to FWorkBook.GetWorksheetCount-1 do begin
    sheet := FWorkBook.GetWorksheetByIndex(i);
    for cell in sheet.Cells do
      if HasFormula(cell) then
        if pos('!', cell^.FormulaValue) > 0 then
          for j:=0 to FWorkbook.GetWorksheetCount-1 do
            if pos(FWorksbook.GetWorksheetByIndex(j).Name, cell1.FormulaValue) = 1 then begin
              extern[n].FirstIndex := j;
              extern[n].LastIndex := j;
              // NOTE: This must be extended to allow a range of sheets !!!
              inc(n);
              if n mod BUF_COUNT = 0 then
                Setlength(extern, Length(extern) + BUF_COUNT);
            end;
  end;

  // Find sheets used in print ranges, repeated cols or repeated rows
  for i:=0 to FWorkbook.GetWorksheetCount-1 do begin
    sheet := FWorkbook.GetWorksheetbyIndex(i);
    with sheet.PageLayout do
      writeIt := (NumPrintRanges > 0) or HasRepeatedCols or HasRepeatedRows;
    if writeIt then begin
      extern[n].FirstIndex := i;
      extern[n].LastIndex := i;
      inc(n);
      if n mod BUF_COUNT = 0 then
        SetLength(extern, Length(extern) + BUF_COUNT);
    end;
  end;

  SetLength(extern, n);

  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_EXTERNSHEET, 2 + 6*n);

  { Count of following REF structures }
  AStream.WriteWord(WordToLE(n));

  { REF record for each sheet }
  for i := 0 to n-1 do
  begin
    AStream.WriteWord(0);                    // Index to EXTERNBOOK record, always 0
    AStream.WriteWord(WordToLE(extern[i]));  // Index to first sheet in EXTERNBOOK sheet list
    AStream.WriteWord(WordToLE(extern[i]));  // Index to last sheet in EXTERNBOOK sheet list
  end;
end;        *)

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 FONT record.
  The font data is passed as an instance of TsFont
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteFONT(AStream: TStream; AFont: TsFont);
var
  Len: Byte;
  WideFontName: WideString;
  optn: Word;
begin
  if AFont = nil then  // this happens for FONT4 in case of BIFF
    exit;

  if AFont.FontName = '' then
    raise EFPSpreadsheetWriter.Create('Font name not specified.');
  if AFont.Size <= 0.0 then
    raise EFPSpreadsheetWriter.Create('Font size not specified.');

  WideFontName := UTF8Decode(AFont.FontName);
  Len := Length(WideFontName);

  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_FONT, 16 + Len * Sizeof(WideChar));

  { Height of the font in twips = 1/20 of a point }
  AStream.WriteWord(WordToLE(round(AFont.Size*20)));

  { Option flags }
  optn := 0;
  if fssBold in AFont.Style then optn := optn or $0001;
  if fssItalic in AFont.Style then optn := optn or $0002;
  if fssUnderline in AFont.Style then optn := optn or $0004;
  if fssStrikeout in AFont.Style then optn := optn or $0008;
  AStream.WriteWord(WordToLE(optn));

  { Color index }
  AStream.WriteWord(WordToLE(PaletteIndex(AFont.Color)));

  { Font weight }
  if fssBold in AFont.Style then
    AStream.WriteWord(WordToLE(INT_FONT_WEIGHT_BOLD))
  else
    AStream.WriteWord(WordToLE(INT_FONT_WEIGHT_NORMAL));

  { Escapement type }
  AStream.WriteWord(WordToLE(ord(AFont.Position)));

  { Underline type }
  if fssUnderline in AFont.Style then
    AStream.WriteByte(1)
  else
    AStream.WriteByte(0);

  { Font family }
  AStream.WriteByte(0);

  { Character set }
  AStream.WriteByte(0);

  { Not used }
  AStream.WriteByte(0);

  { Font name: Unicodestring, char count in 1 byte }
  AStream.WriteByte(Len);
  { Widestring flags, 1=regular unicode LE string }
  AStream.WriteByte(1);
  AStream.WriteBuffer(WideStringToLE(WideFontName)[1], Len * Sizeof(WideChar));
end;

{@@ ----------------------------------------------------------------------------
  Writes the Excel 8 FONT records needed for the fonts used in the workbook.
-------------------------------------------------------------------------------}
procedure TsSpreadBiff8Writer.WriteFonts(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to Workbook.GetFontCount-1 do
    WriteFONT(AStream, Workbook.GetFont(i));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 FORMAT record
  ("Format" is to be understood as "number format" here).
-------------------------------------------------------------------------------}
procedure TsSpreadBiff8Writer.WriteFORMAT(AStream: TStream;
  ANumFormatStr: String; ANumFormatIndex: Integer);
type
  TNumFormatRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FormatIndex: Word;
    FormatStringLen: Word;
    FormatStringFlags: Byte;
  end;
var
  len: Integer;
  ws: widestring;
  rec: TNumFormatRecord;
  buf: array of byte;
begin
  ws := UTF8Decode(ANumFormatStr);
  len := Length(ws);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_FORMAT);
  rec.RecordSize := WordToLE(2 + 2 + 1 + len * SizeOf(WideChar));

  { Format index }
  rec.FormatIndex := WordToLE(ANumFormatIndex);

  { Format string }
  { - length of string = 16 bits }
  rec.FormatStringLen := WordToLE(len);
  { - Widestring flags, 1 = regular unicode LE string }
  rec.FormatStringFlags := 1;
  { - Copy the text characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + SizeOf(WideChar)*len);
  Move(rec, buf[0], SizeOf(rec));
  Move(ws[1], buf[SizeOf(rec)], len*SizeOf(WideChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(rec) + SizeOf(WideChar)*len);

  { Clean up }
  SetLength(buf, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 HEADER or FOOTER record, depending on AIsHeader.
  Overridden because of wide string
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteHeaderFooter(AStream: TStream;
  AIsHeader: Boolean);
var
  wideStr: WideString;
  len: Integer;
  id: Word;
begin
  with FWorksheet.PageLayout do
    if AIsHeader then
    begin
      if (Headers[HEADER_FOOTER_INDEX_ALL] = '') then
        exit;
      wideStr := UTF8Decode(Headers[HEADER_FOOTER_INDEX_ALL]);
      id := INT_EXCEL_ID_HEADER;
    end else
    begin
      if (Footers[HEADER_FOOTER_INDEX_ALL] = '') then
        exit;
      wideStr := UTF8Decode(Footers[HEADER_FOOTER_INDEX_ALL]);
      id := INT_EXCEL_ID_FOOTER;
    end;
  len := Length(wideStr);

  { BIFF record header }
  WriteBiffHeader(AStream, id, 3 + len*sizeOf(wideChar));

  { 16-bit string length }
  AStream.WriteWord(WordToLE(len));

  { Widestring flags, 1=regular unicode LE string }
  AStream.WriteByte(1);

  { Characters }
  AStream.WriteBuffer(WideStringToLE(wideStr)[1], len * SizeOf(WideChar));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 HYPERLINK record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteHyperlink(AStream: TStream;
  AHyperlink: PsHyperlink; AWorksheet: TsWorksheet);
var
  temp: TStream;
  guid: TGUID;
  widestr: widestring;
  ansistr: ansistring;
  descr: String;
  fn: String;
  flags: DWord;
  size: Integer;
  cell: PCell;
  target, bookmark: String;
  u: TUri;
  isInternal: Boolean;
  dirUpCounter: Integer;
begin
  cell := AWorksheet.FindCell(AHyperlink^.Row, AHyperlink^.Col);
  if (cell = nil) or (AHyperlink^.Target='') then
    exit;

  descr := AWorksheet.ReadAsText(cell);      // Hyperlink description
  SplitHyperlink(AHyperlink^.Target, target, bookmark);
  u := ParseURI(AHyperlink^.Target);
  isInternal := (target = '') and (bookmark <> '');
  fn := '';        // Name of local file
  if target <> '' then
  begin
    if (u.Protocol='') then
      fn := target
    else
      UriToFileName(target, fn);
    ForcePathDelims(fn);
  end;

  // Since the length of the record is not known in the first place we write
  // the data to a temporary stream at first.
  temp := TMemoryStream.Create;
  try
    { Cell range using the same hyperlink - we support only single cells }
    temp.WriteWord(WordToLE(cell^.Row));   // first row
    temp.WriteWord(WordToLE(cell^.Row));   // last row
    temp.WriteWord(WordToLE(cell^.Col));   // first column
    temp.WriteWord(WordToLE(cell^.Col));   // last column

    { GUID of standard link }
    guid := StringToGuid('{79EAC9D0-BAF9-11CE-8C82-00AA004BA90B}');
    temp.WriteBuffer(guid, SizeOf(guid));

    { unknown }
    temp.WriteDWord(DWordToLe($00000002));

    { option flags }
    flags := 0;
    if isInternal then
      flags := MASK_HLINK_TEXTMARK or MASK_HLINK_DESCRIPTION
    else
      flags := MASK_HLINK_LINK;
    if SameText(u.Protocol, 'file') or SameText(u.Protocol, 'http') or SameText(u.Protocol, 'ftp') then
      flags := flags or MASK_HLINK_ABSOLUTE;
    if descr <> AHyperlink^.Target then
      flags := flags or MASK_HLINK_DESCRIPTION;  // has description
    if bookmark <> '' then
      flags := flags or MASK_HLINK_TEXTMARK;     // link contains a bookmark
    temp.WriteDWord(DWordToLE(flags));

    { description }
    if flags and MASK_HLINK_DESCRIPTION <> 0 then
    begin
      widestr := UTF8Decode(descr);
      { Character count incl trailing zero }
      temp.WriteDWord(DWordToLE(Length(wideStr) + 1));
      { Character array (16-bit characters), plus trailing zeros }
      temp.WriteBuffer(wideStr[1], (Length(wideStr)+1)*SizeOf(widechar));
    end;

    if target <> '' then
    begin
      if (fn <> '') then  // URI is a local file
      begin
       { GUID of file moniker }
        guid := StringToGuid('{00000303-0000-0000-C000-000000000046}');
        temp.WriteBuffer(guid, SizeOf(guid));
        { Convert to ansi - should be DOS 8.3, but this is not necessary }
        ansistr := UTF8ToAnsi(fn);
        { Directory-up level counter  }
        dirUpCounter := 0;
        if not FileNameIsAbsolute(ansistr) then
          while (pos ('..' + PathDelim, ansistr) = 1) do
          begin
            inc(dirUpCounter);
            Delete(ansistr, 1, Length('..'+PathDelim));
          end;
        temp.WriteWord(WordToLE(dirUpCounter));
        { Character count of file name incl trailing zero }
        temp.WriteDWord(DWordToLe(Length(ansistr)+1));
        { Character array of file name (8-bit characters), plus trailing zero }
        temp.WriteBuffer(ansistr[1], Length(ansistr)+1);
        { Unknown }
        temp.WriteDWord(DWordToLE($DEADFFFF));
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        temp.WriteDWord(0);
        { Size of following file link fields }
        widestr := UTF8ToUTF16(fn);
        size := 4 + 2 + Length(wideStr)*SizeOf(widechar);
        temp.WriteDWord(DWordToLE(size));
        if size > 0 then
        begin
          { Character count of extended file name }
          temp.WriteDWord(DWordToLE(Length(widestr)*SizeOf(WideChar)));
          { Unknown }
          temp.WriteWord(WordToLE($0003));
          { Character array, 16-bit characters, NOT ZERO-TERMINATED! }
          temp.WriteBuffer(widestr[1], Length(wideStr)*SizeOf(WideChar));
        end;
      end
      else begin  { Hyperlink target is a URL }
        widestr := UTF8Decode(target);
        { GUID of URL Moniker }
        guid := StringToGUID('{79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}');
        temp.WriteBuffer(guid, SizeOf(guid));
        { Character count incl trailing zero }
        temp.WriteDWord(DWordToLE(Length(wideStr)+1)*SizeOf(wideChar));
        { Character array plus trailing zero (16-bit characters), plus trailing zeros }
        temp.WriteBuffer(wideStr[1], (length(wideStr)+1)*SizeOf(wideChar));
      end;
    end; // hkURI

    // Hyperlink contains a text mark (#)
    if bookmark <> '' then
    begin
      // Convert to 16-bit characters
      widestr := UTF8Decode(bookmark);
      { Character count of text mark, incl trailing zero }
      temp.WriteDWord(DWordToLE(Length(wideStr) + 1));
      { Character array (16-bit characters) plus trailing zeros }
      temp.WriteBuffer(wideStr[1], (Length(wideStr)+1) * SizeOf(WideChar));
    end;

    { BIFF record header }
    WriteBIFFHeader(AStream, INT_EXCEL_ID_HYPERLINK, temp.Size);

    { Record data }
    temp.Position := 0;
    AStream.CopyFrom(temp, temp.Size);

  finally
    temp.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes all hyperlinks
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteHyperlinks(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  hyperlink: PsHyperlink;
begin
  for hyperlink in AWorksheet.Hyperlinks do begin
    { Write HYPERLINK record }
    WriteHyperlink(AStream, hyperlink, AWorksheet);
    { Write HYPERLINK TOOLTIP record }
    if hyperlink^.Tooltip <> '' then
      WriteHyperlinkTooltip(AStream, hyperlink^.Row, hyperlink^.Col, hyperlink^.Tooltip);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a HYPERLINK TOOLTIP record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteHyperlinkTooltip(AStream: TStream;
  const ARow, ACol: Cardinal; const ATooltip: String);
var
  widestr: widestring;
begin
  widestr := UTF8Decode(ATooltip);

  { BIFF record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_HLINKTOOLTIP,
    10 + (Length(wideStr)+1) * SizeOf(widechar));

  { Repeated record ID }
  AStream.WriteWord(WordToLe(INT_EXCEL_ID_HLINKTOOLTIP));

  { Cell range using the same hyperlink tooltip - we support only single cells }
  AStream.WriteWord(WordToLE(ARow));   // first row
  AStream.WriteWord(WordToLE(ARow));   // last row
  AStream.WriteWord(WordToLE(ACol));   // first column
  AStream.WriteWord(WordToLE(ACol));   // last column

  { Tooltop characters, no length, but trailing zero }
  AStream.WriteBuffer(wideStr[1], (Length(widestr)+1)*SizeOf(wideChar));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 INDEX record

    nm = (rl - rf - 1) / 32 + 1 (using integer division)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteINDEX(AStream: TStream);
begin
  { BIFF Record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_INDEX, 16);

  { Not used }
  AStream.WriteDWord(DWordToLE(0));

  { Index to first used row, rf, 0 based }
  AStream.WriteDWord(DWordToLE(0));

  { Index to first row of unused tail of sheet, rl, last used row + 1, 0 based }
  AStream.WriteDWord(DWordToLE(0));

  { Absolute stream position of the DEFCOLWIDTH record of the current sheet.
    If it doesn't exist, the offset points to where it would occur. }
  AStream.WriteDWord(DWordToLE($00));

  { Array of nm absolute stream positions of the DBCELL record of each Row Block }

  { OBS: It seems to be no problem just ignoring this part of the record }
end;

{@@ ----------------------------------------------------------------------------
  Depending on the presence of Rich-text formatting information in the cell
  record, writes an Excel 8 LABEL record (string cell value only), or
  RSTRING record (string cell value + rich-text formatting runs).

  If the cell text (incl its rich-text formatting) is found in the
  SharedStringtable then LABELSST record is written.

  If the string length exceeds 32758 bytes, the string will be truncated,
  a note will be left in the workbooks log.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteLABEL(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: String; ACell: PCell);
var
  L: Word;
  WideStr: WideString;
  rec: TBIFF8_LabelRecord;
  recSST: TBIFF8_LabelSSTRecord;
  buf: array of byte;
  i, nRuns, idx: Integer;
  rtfRuns: TBiff8_RichTextFormattingRuns;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  // If string is in SST write a LABELSST record
  idx := IndexOfSharedString(ACell^.UTF8StringValue, ACell^.RichTextParams);
  if idx > -1 then begin
    recSST.RecordID := WordToLE(INT_EXCEL_ID_LABELSST);
    recSST.RecordSize := WordToLE(SizeOf(recSST) - SizeOf(TsBiffHeader));
    recSST.Row := WordToLE(ARow);
    recSST.Col := WordToLE(ACol);
    recSST.XFIndex := WordToLE(FindXFIndex(ACell^.FormatIndex));
    recSST.SSTIndex := DWordToLE(idx);
    AStream.Write(recSST, SizeOf(recSST));
    exit;
  end;

  // If string is not in SST write a standard LABEL cell
  WideStr := UTF8Decode(FixLineEnding(AValue)); //to UTF16
  if WideStr = '' then begin
    // Badly formatted UTF8String (maybe ANSI?)
    if Length(AValue)<>0 then begin
      //Quite sure it was an ANSI string written as UTF8, so raise exception.
      raise EFPSpreadsheetWriter.CreateFmt(rsUTF8TextExpectedButANSIFoundInCell, [GetCellString(ARow, ACol)]);
    end;
    Exit;
  end;

  if Length(WideStr) > FLimitations.MaxCharsInTextCell then begin
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    SetLength(WideStr, FLimitations.MaxCharsInTextCell); //may corrupt the string (e.g. in surrogate pairs), but... too bad.
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
      FLimitations.MaxCharsInTextCell, GetCellString(ARow, ACol)
    ]);
  end;
  L := Length(WideStr);
  nRuns := Length(ACell^.RichTextParams);

  { BIFF record header }
  rec.RecordID := WordToLE(IfThen(nRuns > 0, INT_EXCEL_ID_RSTRING, INT_EXCEL_ID_LABEL));
  (*
  recSize := SizeOf(TBiff8_LabelRecord) - SizeOf(TsBiffHeader) + L*SizeOf(WideChar);
  if nRuns > 0 then
    inc(recSize, SizeOf(Word) + nRunms * SizeOf(TBiff8_RichTextFormattingRun);
  if n
    *)
  rec.RecordSize := SizeOf(TBiff8_LabelRecord) - SizeOf(TsBiffHeader) + L *SizeOf(WideChar);
  if nRuns > 0 then
    inc(rec.RecordSize, SizeOf(Word) + nRuns * SizeOf(TBiff8_RichTextFormattingRun));
  rec.RecordSize := WordToLE(rec.RecordSize);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { Index to XF record, according to formatting }
  rec.XFIndex := WordToLE(FindXFIndex(ACell^.FormatIndex));

  { Byte String with 16-bit length }
  rec.TextLen := WordToLE(L);

  { Byte flags }
  rec.TextFlags := 1;                     // means regular unicode LE encoding
  // Excel does not write the Rich-Text flag probably because rich-text info
  // is located differently in the RSTRING record.

  { Copy the text characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + L*SizeOf(WideChar));
  Move(rec, buf[0], SizeOf(Rec));
  Move(WideStringToLE(WideStr)[1], buf[SizeOf(Rec)], L*SizeOf(WideChar));

  { Write out buffer }
  AStream.WriteBuffer(buf[0], SizeOf(rec) + L*SizeOf(WideChar));

  { Write rich-text information in case of RSTRING record }
  if (nRuns > 0) then
  begin
    { Write number of rich-text formatting runs }
    AStream.WriteWord(WordToLE(nRuns));

    { Write array of rich-text formatting runs }
    SetLength(rtfRuns, nRuns);
    for i:=0 to nRuns-1 do
    begin
      // index of first character of formatted part, 0-based in file, 1-based in fps
      rtfRuns[i].FirstIndex := WordToLE(ACell^.RichTextParams[i].FirstIndex - 1);
      // Index of new font. Be aware of font #4 missing in BIFF!
      if ACell^.RichTextParams[i].FontIndex >= 4 then
        rtfRuns[i].FontIndex := WordToLE(ACell^.RichTextParams[i].FontIndex + 1) else
        rtfRuns[i].FontIndex := WordToLE(ACell^.RichTextParams[i].FontIndex);
    end;
    AStream.WriteBuffer(rtfRuns[0], nRuns * SizeOf(TBiff8_RichTextFormattingRun));
  end;

  { Clean up }
  SetLength(rtfRuns, 0);
  SetLength(buf, 0);
end;

procedure TsSpreadBIFF8Writer.WriteMergedCells(AStream: TStream;
  AWorksheet: TsWorksheet);
const
  MAX_PER_RECORD = 1026;
var
  n0, n: Integer;
  rng: PsCellRange;
  newRecord: Boolean;
begin
  n0 := AWorksheet.MergedCells.Count;
  n := Min(n0, MAX_PER_RECORD);
  newRecord := true;
  for rng in AWorksheet.MergedCells do
  begin
    if newRecord then
    begin
      newRecord := false;
      { BIFF record header }
      WriteBIFFHeader(AStream, INT_EXCEL_ID_MERGEDCELLS, 2 + n*8);
      { Number of cell ranges in this record }
      AStream.WriteWord(WordToLE(n));
    end;
    { Write range data }
    AStream.WriteWord(WordToLE(rng^.Row1));
    AStream.WriteWord(WordToLE(rng^.Row2));
    AStream.WriteWord(WordToLE(rng^.Col1));
    AStream.WriteWord(WordToLE(rng^.Col2));

    dec(n);
    if n = 0 then begin
      newRecord := true;
      dec(n0, MAX_PER_RECORD);
      n := Min(n0, MAX_PER_RECORD);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the first MSODRAWING record to file. It is needed for a comment
  attached to a cell, but also for embedded shapes (currently not supported).

<pre>
  Structure of this record:
                                                    Type    Ver   Inst
     Dg container                                   $F002          0
       |--- FDG record                              $F008    0     1
       |--- SpGr container                          $F003          0
             |---- Sp container  (group shape)      $F004          0
             |       |---- FSpGr record             $F009    1     0
MSODRAWING1  |       |---- FSp record               $F00A    2     0
................................................................................
MSODRAWING2  |---- Sp container  (child shape)      $F004          0
                     |---- FSp record               $F00A    2    202 (Textbox)
                     |---- FOpt record              $F00B    3    13 (num props)
                     |---- Client anchor record     $F010    0     0
                     |---- Client data record       $F011    0     0
</pre>
-------------------------------------------------------------------------------}
procedure TsSpreadBiff8Writer.WriteMSODrawing1(AStream: TStream; ANumShapes: Word;
  AComment: PsComment);
const
  DRAWING_ID = 1;
var
  len: DWord;
  tmpStream: TMemoryStream;
begin
  tmpStream := TMemoryStream.Create;
  try
    { OfficeArtDgContainer record (container of drawing) }
    len := 224 + 152*(ANumShapes - 1);
    WriteMSODgContainer(tmpStream, len);

    { OfficeArtFdg record (info on shapes: num shapes, drawing ID, last Obj ID ) }
    WriteMSOFdgRecord(tmpStream, ANumShapes + 1, DRAWING_ID, SHAPEID_BASE + ANumShapes);

    { OfficeArtSpGrContainer record (shape group container) }
    len := 200 + 152*(ANumShapes - 1);
    WriteMSOSpGrContainer(tmpStream, len);

    { OfficeArtSpContainer record }
    WriteMSOSpContainer(tmpStream, 40);

    { OfficeArtFSpGr record }
    WriteMSOFSpGrRecord(tmpStream, 0, 0, 0, 0);       // 16 + 8 bytes

    { OfficeArtFSp record }
    WriteMSOFSpRecord(tmpStream, SHAPEID_BASE, MSO_SPT_NOTPRIMITIVE,
      MSO_FSP_BITS_GROUP + MSO_FSP_BITS_PATRIARCH);   // 8 + 8 bytes

    { Data for the 1st comment }
    WriteMSODrawing2_Data(tmpStream, AComment, SHAPEID_BASE + 1);

    { Write the BIFF stream }
    tmpStream.Position := 0;
    len := tmpStream.Size;
    WriteBiffHeader(AStream, INT_EXCEL_ID_MSODRAWING, tmpStream.Size);
    AStream.CopyFrom(tmpStream, tmpStream.Size);
  finally
    tmpStream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the MSODRAWING record which occurs before the OBJ record.
  Not to be used for the very first OBJ record where the record must be
  WriteMSODrawing1 + WriteMSODrawing2_Data
-------------------------------------------------------------------------------}
procedure TsSpreadBiff8Writer.WriteMSODrawing2(AStream: TStream;
  AComment: PsComment; AObjID: Word);
var
  tmpStream: TStream;
  len: Word;
begin
  tmpStream := TMemoryStream.Create;
  try
    { Shape data for cell comment }
    WriteMSODrawing2_Data(tmpStream, AComment, SHAPEID_BASE + AObjID);

    { Get size of data stream }
    len := tmpStream.Size;

    { BIFF Header }
    WriteBiffHeader(AStream, INT_EXCEL_ID_MSODRAWING, len);

    { Copy MSO data to BIFF stream }
    tmpStream.Position := 0;
    AStream.CopyFrom(tmpStream, len);
  finally
    tmpStream.Free;
  end;
end;

procedure TsSpreadBiff8Writer.WriteMSODrawing2_Data(AStream: TStream;
  AComment: PsComment; AShapeID: Word);
var
  tmpStream: TStream;
  len: Cardinal;
begin
  // We write all the record data to a temporary stream to get the record
  // size (it depends on the number of properties written to the FOPT record.
  // The record size is needed in the very first SpContainer record...

  tmpStream := TMemoryStream.Create;
  try
    { OfficeArtFSp record }
    WriteMSOFSpRecord(tmpStream, AShapeID, MSO_SPT_TEXTBOX,
      MSO_FSP_BITS_HASANCHOR + MSO_FSP_BITS_HASSHAPETYPE);

    { OfficeArtFOpt record }
    WriteMSOFOptRecord_Comment(tmpStream);

    { OfficeArtClientAnchor record }
    WriteMSOClientAnchorSheetRecord(tmpStream,
      AComment^.Row + 1, AComment^.Col + 1, AComment^.Row + 3, AComment^.Col + 5,
      691, 486, 38, 26,
      true, true
    );

    { OfficeArtClientData record }
    WriteMSOClientDataRecord(tmpStream);

    // Now we know the record size
    len := tmpStream.Size;

    // Write an OfficeArtSpContainer to the stream provided...
    WriteMSOSpContainer(AStream, len+8);    // !!! for some reason, Excel wants here 8 additional bytes !!!

    // ... and append the data from the temporary stream.
    tmpStream.Position := 0;
    AStream.Copyfrom(tmpStream, len);

  finally
    tmpStream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the MSODRAWING record which must occur immediately before a TXO record
-------------------------------------------------------------------------------}
procedure TsSpreadBiff8Writer.WriteMSODRAWING3(AStream: TStream);
begin
  { BIFF Header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_MSODRAWING, 8);

  { OfficeArtClientTextbox record: Text-related data for a shape }
  WriteMSOClientTextBoxRecord(AStream);
end;

{@@ ----------------------------------------------------------------------------
  Writes a NOTE record for a comment attached to a cell
-------------------------------------------------------------------------------}
procedure TsSpreadBiff8Writer.WriteNOTE(AStream: TStream; AComment: PsComment;
  AObjID: Word);
const
  AUTHOR: ansistring = 'author';
var
  len: Integer;
begin
  len := Length(AUTHOR) * sizeOf(ansichar);

  { BIFF Header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_NOTE));  // ID of NOTE record
  AStream.WriteWord(WordToLE(12+len));             // Size of NOTE record

  { Record data }
  AStream.WriteWord(WordToLE(AComment^.Row));      // Row index of cell
  AStream.WriteWord(WordToLE(AComment^.Col));      // Column index of cell
  AStream.WriteWord(0);                            // Flags
  AStream.WriteWord(WordToLE(AObjID));             // Object identifier (1, ...)
  AStream.WriteWord(len);                          // Char length of author string
  AStream.WriteByte(0);                            // Flag for 8-bit characters
  AStream.WriteBuffer(AUTHOR[1], len);             // Author
  AStream.WriteByte(0);                            // Unused
end;

{@@ ----------------------------------------------------------------------------
  Writes an OBJ record - belongs to the records required for cell comments
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteOBJ(AStream: TStream; AObjID: Word);
var
  guid: TGuid;
begin
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_OBJ));
  AStream.WriteWord(WordToLE(52));

  AStream.WriteWord(WordToLE($0015));  // Subrecord ftCmo
  AStream.WriteWord(WordToLE(18));     // Subrecord size: 18 bytes
  AStream.WriteWord(WordToLE($0019));  // Object type: Comment
  AStream.WriteWord(WordToLE(AObjID)); // Object ID number (1, ... )
  AStream.WriteWord(WordToLE($4011));  // Option flags automatic line style, locked when sheet is protected
  AStream.WriteDWord(0);               // Unused
  AStream.WriteDWord(0);               // Unused
  AStream.WriteDWord(0);               // Unused

  AStream.WriteWord(WordToLE($000D));  // Subrecord ftNts
  AStream.WriteWord(WordToLE(22));     // Size of subrecord: 22 bytes
//  CreateGUID(guid);
  FillChar(guid{%H-}, SizeOf(guid), 0);
  AStream.WriteBuffer(guid, 16);       // GUID of comment
  AStream.WriteWord(WordToLE(0));      // shared note (0 = false)
  AStream.WriteDWord(0);               // unused

  AStream.WriteWord(WordToLE($0000));  // Subrecord ftEnd
  AStream.WriteWord(0);                // Size of subrecord: 0 bytes
end;

{@@ ----------------------------------------------------------------------------
  Writes the address of a cell as used in an RPN formula and returns the
  number of bytes written.
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.WriteRPNCellAddress(AStream: TStream;
  ARow, ACol: Cardinal; AFlags: TsRelFlags): Word;
var
  c: Cardinal;       // column index with encoded relative/absolute address info
begin
  AStream.WriteWord(WordToLE(ARow));
  c := ACol and MASK_EXCEL_COL_BITS_BIFF8;
  if (rfRelRow in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW_BIFF8;
  if (rfRelCol in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL_BIFF8;
  AStream.WriteWord(WordToLE(c));
  Result := 4;
end;

function TsSpreadBIFF8Writer.WriteRPNCellAddress3D(AStream: TStream;
  ASheet, ARow, ACol: Cardinal; AFlags: TsRelFlags): Word;
var
  c: Cardinal;
begin
  // Next line is a simplification: We should write the index of the sheet
  // in the REF record here, but these are arranged in the same order as the
  // sheets. --> MUST BE RE-DONE ONCE SHEET RANGES ARE ALLOWED.
  AStream.WriteWord(WordToLE(ASheet));

  // The row/col address is written in relative notation!
  Result := 2 + WriteRPNCellAddress(AStream, ARow, ACol, [rfRelRow, rfRelCol]);
end;


{@@ ----------------------------------------------------------------------------
  Writes row and column offset needed in RPN formulas (unsigned integers!)
  Valid for BIFF2-BIFF5.
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.WriteRPNCellOffset(AStream: TStream;
  ARowOffset, AColOffset: Integer; AFlags: TsRelFlags): Word;
var
  c: Word;
  r: SmallInt;
begin
  // row address
  r := SmallInt(ARowOffset);
  AStream.WriteWord(WordToLE(Word(r)));

  // Encoded column address
  c := word(AColOffset) and MASK_EXCEL_COL_BITS_BIFF8;
  if (rfRelRow in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW_BIFF8;
  if (rfRelCol in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL_BIFF8;
  AStream.WriteWord(WordToLE(c));

  Result := 4;
end;

{@@ ----------------------------------------------------------------------------
  Writes the address of a cell range as used in an RPN formula and returns the
  count of bytes written.
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.WriteRPNCellRangeAddress(AStream: TStream;
  ARow1, ACol1, ARow2, ACol2: Cardinal; AFlags: TsRelFlags): Word;
var
  c: Cardinal;       // column index with encoded relative/absolute address info
begin
  AStream.WriteWord(WordToLE(ARow1));
  AStream.WriteWord(WordToLE(ARow2));

  c := ACol1;
  if (rfRelCol in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL;
  if (rfRelRow in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW;
  AStream.WriteWord(WordToLE(c));

  c := ACol2;
  if (rfRelCol2 in AFlags) then c := c or MASK_EXCEL_RELATIVE_COL;
  if (rfRelRow2 in AFlags) then c := c or MASK_EXCEL_RELATIVE_ROW;
  AStream.WriteWord(WordToLE(c));

  Result := 8;
end;

{ Writes a buffer containing a string (with header) and its associated
  rich-text parameters to an EXCEL record.
  Since the size in an EXCEL8 record is limited to 8224 bytes
  ABytesAvail specifies how many bytes are available. If the buffer is longer
  then the string is split at an appropriate position.
  The procedure also handles the case that the remaining bytes of the string/
  rich-text are written in a following CONTINUE record.
  The number of written bytes is returned in the corresponding parameter.
  If a split has occured in the string part of the buffer the following CONTINUE
  record must begin with an Option byte; this is signaled by AContinueInString.
  The function returns TRUE if the entire buffer has been written completely,
  otherweise CONTINUE records are required. }
function TsSpreadBIFF8Writer.WriteRichTextStream(AStream: TStream;
  ABuffer: TMemoryStream; ABytesAvail: Integer; out ABytesWritten: Integer;
  out AContinueInString: Boolean): Boolean;
var
  strLen: Word;
  hasRtp: Boolean;
  nRtp: Integer;
  bytesToWrite: Integer;
  savedStartPos: Integer;
  strSize, rtpSize, hdrSize: Int64;
begin
  Result := false;

  ABytesWritten := 0;
  AContinueInString := false;

  // Read string header and get string length and count of rich-text parameters
  savedStartPos := ABuffer.Position;
  ABuffer.Position := 0;
  strLen := LEToN(ABuffer.ReadWord);      // string length (character count)
  hasRtp := ABuffer.ReadByte and 4 <> 0;  // Rich-text params available?
  hdrSize := SizeOf(strlen) + SizeOf(byte);
  if hasRtp then
  begin
    nRtp := LEToN(ABuffer.ReadWord);   // number of rich-text formatting runs
    inc(hdrSize, SizeOf(word));
    rtpSize := Int64(nRtp) * 4;  // 4 bytes per rich-text formatting run
  end else
  begin
    nRtp := 0;
    rtpSize := 0;
  end;
  strSize := Int64(strLen) * SizeOf(WideChar);  // String length in bytes

  // Begin writing
  ABuffer.Position := savedStartPos;

  // Case 1: If the function has been called for the 1st time,
  // we must write the string header
  if ABuffer.Position = 0 then
  begin
    // Keep header plus 1st character together - they must not be separated
    bytesToWrite := hdrSize + SizeOf(WideChar);
    if ABytesAvail < bytesToWrite then
    begin
      // not enough memory left
      AContinueInString := true;  // The following CONTINUE record must begin with the Options byte
      exit;
    end;
    AStream.CopyFrom(ABuffer, bytesToWrite);
    inc(ABytesWritten, bytesToWrite);
  end;

  // Case 2; Here some part of the string already has been written, and the
  // buffer stream is somewhere in the string part
  if ABuffer.Position < Int64(hdrSize) + Int64(strSize) - 1 then
  begin
    bytesToWrite := hdrSize + strSize - ABuffer.Position;
    if bytesToWrite > ABytesAvail then
    begin
      bytesToWrite := ABytesAvail;
      // Make sure to split between widechars
      if odd(bytesToWrite) then dec(bytesToWrite);
      AStream.CopyFrom(ABuffer, bytesToWrite);
      inc(ABytesWritten, bytesToWrite);
      AContinueInString := true;
      exit;
    end;
    AStream.CopyFrom(ABuffer, bytesToWrite);
    inc(ABytesWritten, bytesToWrite);
  end;

  // Case 3: The string has been written fully, but the buffer stream is somewhere
  // in the rich-text formatting runs
  if hasRtp and (ABuffer.Position >= hdrSize + strSize) then
  begin
    bytesToWrite := hdrSize + strSize + rtpSize - ABuffer.Position;
    if bytesToWrite > ABytesAvail then
    begin
      // Make sure to split between rich-text formatting runs. Each run is 4 bytes.
      bytesToWrite := (ABytesAvail div 4) * 4;
      AStream.CopyFrom(ABuffer, BytesToWrite);
      inc(ABytesWritten, bytesToWrite);
      exit;
    end;
    AStream.CopyFrom(ABuffer, BytesToWrite);
    inc(ABytesWritten, bytesToWrite);
  end;

  // If the procedure gets to this point the buffer has been written completely.
  Result := true;
end;

{@@ ----------------------------------------------------------------------------
  Writes the SharedStringTable (SST) to the stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteSST(AStream: TStream);

  function Is8BitString(s: String): boolean;
  var
    i: Integer;
  begin
    Result := false;
    for i:=1 to Length(s) do
      if s[i] > #127 then exit;
    Result := true;
  end;

var
  sizePos: Int64;
  bytesWritten, totalBytesWritten: Integer;
  i, j: Integer;
  rtParams: TsRichTextParams;
  bytesAvail: Integer;
  isASCII: Boolean;
  textIndex, rtIndex: Integer;
  complete: Boolean;
  flag: Byte;
  ws: Widestring;
  s: String;
  rs: RawByteString;

begin
  if FSharedStringTable.Count = 0 then
    exit;

  { Write BIFF header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_SST));
  sizePos := AStream.Position;
  AStream.WriteWord(0);     // Size of record will be written later when we know it

  { Number of strings in workbook }
  AStream.WriteDWord(DWordToLE(FNumStrings));

  { Number of strings in SST }
  AStream.WriteDWord(DWordToLE(FSharedStringTable.Count));

  { Now begins writing of strings. Take care of overflow into following
    CONTINUE records if the maximum record size (MAX_BYTES_IN_RECORD) is
    exceeded. }
  totalBytesWritten := 8;
  for i:=0 to FSharedStringTable.Count-1 do
  begin
    s := FixLineEnding(FSharedStringTable.Strings[i]);
    isASCII := Is8BitString(s);
    if isASCII then
    begin
      rs := s;
      if Length(s) > FLimitations.MaxCharsInTextCell then
      begin
        SetLength(rs, FLimitations.MaxCharsInTextCell);
        FWorkbook.AddErrorMsg(rsTruncateTooLongText, [FLimitations.MaxCharsInTextCell]);
      end;
    end else
    begin
      ws := WideStringToLE(UTF8ToUTF16(s));
      SetLength(rs, Length(ws) * SizeOf(widechar));
      Move(ws[1], rs[1], Length(rs));
      if Length(ws) > FLimitations.MaxCharsInTextCell then
      begin
        SetLength(ws, FLimitations.MaxCharsInTextCell);
        FWorkbook.AddErrorMsg(rsTruncateTooLongText, [FLimitations.MaxCharsInTextCell]);
      end;
    end;

    SetLength(rtParams, Length(TsRichTextParams(FSharedStringTable.Objects[i])));
    for j := 0 to High(rtParams) do begin
      rtParams[j] := TsRichTextParams(FSharedStringTable.Objects[i])[j];
      // Index of new font. Be aware of font #4 missing in BIFF!
      if rtParams[j].FontIndex >= 4 then inc(rtParams[j].FontIndex);
    end;

    textIndex := 1;
    rtIndex := 0;
    repeat
      bytesAvail := MAX_BYTES_IN_RECORD - totalBytesWritten;
      bytesWritten := WriteStringHelper(AStream, rs, rtParams, isASCII,
        bytesAvail, textIndex, rtIndex, complete);
      inc(totalBytesWritten, bytesWritten);
      dec(bytesAvail, bytesWritten);

      // String is not complete --> we need a CONTINUE record
      if not complete then begin
        FixRecordSize(AStream, sizePos, totalBytesWritten);
        BeginCONTINUERecord(AStream, sizePos);
        if (textIndex <> -1) and (textIndex <> 1) then begin
          // Text is split: the string flag must be repeated
          flag := IfThen(IsASCII, 0, 1);
          AStream.WriteByte(flag);
          totalBytesWritten := 1;
        end else
          totalBytesWritten := 0;
      end;
    until complete;
  end;

  // Write size word of the current record
  FixRecordSize(AStream, sizePos, totalBytesWritten);
end;


(*
procedure TsSpreadBIFF8Writer.WriteSST(AStream: TStream);
type
  TBiff8RichTextParam = packed record
    FirstIndex: Word;
    FontIndex: Word;
  end;
  TBiff8RichTextParams = array of TBiff8RichTextParam;
var
  i, j: Integer;
  pSize: Int64;
  s: string;
  ws: WideString;
  rtParams: TsRichTextParams;
  biffRtParams: TBiff8RichTextParams;
  bytesAvail, bytesToWrite, bytesWritten, totalBytesWritten: Integer;
  hasRtp: Boolean;
  hdrSize: Integer;
  flags: Byte;
  startIndex: Integer;
  needCONTINUE: Boolean;

  procedure EndRecord;
  var
    p: Int64;
  begin
    p := AStream.Position;
    AStream.Position := pSize;
    AStream.WriteWord(WordToLE(totalBytesWritten));
    AStream.Position := p;
  end;

  procedure BeginCONTINUERecord;
  begin
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_CONTINUE));
    pSize := AStream.Position;
    AStream.WriteWord(0);
  end;

begin
  if FSharedStringTable.Count = 0 then
    exit;

  { Write BIFF header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_SST));
  pSize := AStream.Position;
  AStream.WriteWord(0);  // Size of record - will be written later

  { Number of strings in workbook }
  AStream.WriteDWord(DWordToLE(FNumStrings));

  { Number of strings in SST }
  AStream.WriteDWord(DWordToLE(FSharedStringTable.Count));

  { Here the strings plus rich-text parameters are following. This is a bit
    complicated because usually there are many strings, but each record can
    hold only 8224 bytes (MAX_BYTES_IN_RECORD) which requires additional
    CONTINUE records. }

  totalBytesWritten := 8;

  for i:=0 to FSharedStringTable.Count-1 do
  begin
    // Assemble the string to be written in a buffer stream
    s := FixLineEnding(FSharedStringTable.Strings[i]);
    ws := WideStringToLE(UTF8Decode(s));
    rtParams := TsRichTextParams(FSharedStringTable.Objects[i]);
    SetLength(biffRtParams, Length(rtParams));
    for j := 0 to High(biffRtParams) do begin
      biffRtParams[j].FirstIndex := WordToLE(rtParams[j].FirstIndex) - 1;
        // character index is 0-based in file, but 1-based in fps.
      biffRtParams[j].FontIndex := WordToLE(rtParams[j].FontIndex);
    end;
    hasRtp := Length(rtParams) > 0;
    hdrsize := IfThen(hasRtp, 3+2, 3);

    bytesAvail := MAX_BYTES_IN_RECORD - totalBytesWritten;

    // (1) String header
    // String header plus 1st character do not fit into current record
    // ---> move everything to a CONTINUE record
    if bytesAvail < hdrsize + SizeOf(WideChar) then begin
      EndRecord;
      BeginCONTINUERecord;     // Begins a CONTINUE record
    end else begin
      { Write string length }
      AStream.WriteWord(WordToLE(Length(ws)));
      { Write string flags byte }
      flags := 1;  // 1 = uncompressed data (= wide chars)
      if hasRtp then inc(flags, 8);  // 8 = has rich-text formatting runs
      inc(totalbytesWritten, 3);
      AStream.Writebyte(flags);
      { Write number of rich-text formatting runs }
      if hasRtp then begin
        AStream.WriteWord(WordToLE(Length(rtParams)));
        inc(totalBytesWritten, 2);
      end;
    end;

    // (2) String characters
    bytesAvail := MAX_BYTES_IN_RECORD - totalBytesWritten;
    if odd(bytesAvail) then dec(bytesAvail);  // Split between widechars
    bytesToWrite := Length(ws) * SizeOf(WideChar);
    needCONTINUE := bytesToWrite > bytesAvail;
    startIndex := 1;
    while needCONTINUE do begin
      // Fill remainder of current record
      bytesWritten := AStream.Write(ws[startIndex], bytesAvail);
      inc(totalBytesWritten, bytesWritten);
      EndRecord;
      BeginCONTINUERecord;
      // Write flag byte because string is split
      AStream.WriteByte(1);
      totalBytesWritten := 1;
      startIndex := StartIndex + bytesWritten div 2;
      bytesAvail := MAX_BYTES_IN_RECORD - totalBytesWritten;
      if odd(bytesAvail) then dec(bytesAvail);
      bytesToWrite := (Length(ws) - startIndex + 1) * SizeOf(WideChar);
      needCONTINUE := bytesToWrite > bytesAvail;
    end;
    if bytesToWrite > 0 then begin
      bytesWritten := AStream.Write(ws[startIndex], bytesToWrite);
      inc(totalBytesWritten, bytesWritten);
    end;

    // (3) Rich-text formatting runs
    bytesAvail := MAX_BYTES_IN_RECORD - totalBytesWritten;
    // Make sure to split between runs
    bytesAvail := (bytesAvail div 4) * 4;
    bytesToWrite := Length(biffRtParams) * 4;  // 4 = size of formatting run
    needCONTINUE := bytesToWrite > bytesAvail;
    startIndex := 0;
    while needCONTINUE do begin
      // Fill remainder of current record
      bytesWritten := AStream.Write(biffRtParams[startIndex], bytesAvail);
      inc(totalBytesWritten, bytesWritten);
      EndRecord;
      BeginCONTINUERecord;
      totalBytesWritten := 0;
      startIndex := startIndex + bytesWritten div 4;
      bytesAvail := MAX_BYTES_IN_RECORD - totalBytesWritten;
      bytesAvail := (bytesAvail div 4) * 4;
      bytesToWrite := (Length(biffRtParams) - startIndex) * 4;
      needCONTINUE := bytesToWrite > bytesAvail;
    end;
    if bytesToWrite > 0 then begin
      bytesWritten := AStream.Write(biffRtParams[startIndex], bytesToWrite);
      inc(totalBytesWritten, bytesWritten);
    end;
  end;

  // Write size word of the current record
  EndRecord;
end;
*)

{@@ ----------------------------------------------------------------------------
  Helper function for writing a string with 8-bit length. Overridden version
  for BIFF8. Called for writing rpn formula string tokens.
  Returns the count of bytes written.
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.WriteString_8BitLen(AStream: TStream;
  AString: String): Integer;
var
  len: Integer;
  wideStr: WideString;
begin
  // string constant is stored as widestring in BIFF8
  wideStr := UTF8Decode(AString);
  len := Length(wideStr);
  AStream.WriteByte(len); // char count in 1 byte
  AStream.WriteByte(1);   // Widestring flags, 1=regular unicode LE string
  AStream.WriteBuffer(WideStringToLE(wideStr)[1], len * Sizeof(WideChar));
  Result := 1 + 1 + len * SizeOf(WideChar);
end;

{@@ ----------------------------------------------------------------------------
  Helper function for writing a string which considers overflow into
  CONTINUE records.

  @@param AText            Text to be written
  @@param ARichTextParams  optional rich-text formatting runs for formatting
                           individual characters
  @@param Is8BitString     if true the string is a compressed 8-bit string,
                           otherwise a widestring
  @@param ABytesAvail      Specifies the number of bytes available in the
                           current record. A BIFF record can hold only
                           MAX_BYTES_IN_RECORD bytes.
  @@param ATextIndex       Index at which writing of the text begins.
  @@param ARichIndex       Index at which writing of the rich-text parameters begins
  @@param AComplete        TRUE if writing is complete, FALSE if text or
                           rich-text parameters do not fit into the current
                           record. In the latter case ATextIndex and ARichIndex
                           are updated for the next call of the method.
                           ATextIndex returns -1, if the text is complete.

  @@return                 Number of bytes written
-------------------------------------------------------------------------------}
function TsSpreadBIFF8Writer.WriteStringHelper(AStream: TStream;
  const AText: RawByteString; const ARichTextParams: TsRichTextParams;
  Is8BitString: Boolean; ABytesAvail: Integer; var ATextIndex, ARichIndex: Integer;
  out AComplete: Boolean): Integer;
const
  COMPRESSED_FLAG: array[boolean] of byte = (1, 0);
var
  hdrSize: Integer;
  bytesToWrite: Integer;
  flags: Byte;
  len: Word;
  nRtp: Integer;
  rtp: TsRichTextParam;
begin
  Result := 0;
  AComplete := false;

  // Text length in characters
  len := Length(AText);
  if not Is8BitString then len := len div 2;
  // Number of rich-text parameters
  nRtp := Length(ARichTextParams);

  // (1) String header
  if (ATextIndex = 1) then
  begin
    hdrSize := IfThen(nRtp = 0, 3, 3+2);

    // String header plus 1st character do not fit into current record
    // ---> the caller must move everything to a CONTINUE record
    if ABytesAvail < hdrSize + SizeOf(WideChar) then
      exit;

    { Write string length (in characters)}
    AStream.WriteWord(WordToLE(len));
    inc(Result, 2);

    { Write string flags byte }
    flags := COMPRESSED_FLAG[Is8BitString];
    if Length(ARichTextParams) > 0 then
      inc(flags, 8);        // 8 = has rich-text formatting runs
    // Note: Asian phonetic not supported here!
    AStream.WriteByte(flags);
    inc(Result, 1);

    { Write number of rich-text formatting runs }
    if Length(ARichTextParams) > 0 then begin
      AStream.WriteWord(WordToLE(nRtp));
      inc(Result, 2);
    end;
  end;

  // (2) String characters
  if ATextIndex <> -1 then begin
    bytesToWrite := Length(AText) - ATextIndex + 1;
    if bytesToWrite > ABytesAvail - Result then begin
      bytesToWrite := ABytesAvail - Result;
      // Make sure to split widestring between widechars
      if not Is8bitString and odd(bytesToWrite) then
        dec(bytesToWrite);
      inc(Result, AStream.Write(AText[ATextIndex], bytesToWrite));
      inc(ATextIndex, bytesToWrite);
      exit;
    end;

    inc(Result, AStream.Write(AText[ATextIndex], bytesToWrite));
    ATextIndex := -1;  // String is complete here
  end;

  // (3) Rich-text formatting runs
  if nRtp = 0 then begin
    AComplete := true;
    exit;
  end;

  while (ARichIndex < nRtp) and (ABytesAvail - Result >= 4) do
  begin
    rtp := ARichTextParams[ARichIndex];
    if rtp.FirstIndex > len then begin
      ARichIndex := MaxInt;
      break;
    end;
    // Make sure to split between runs
    AStream.WriteWord(WordToLE(rtp.FirstIndex - 1));
      // character index is 0-based in file, but 1-based in fps
    AStream.WriteWord(WordtoLE(rtp.FontIndex));
    inc(Result, 4);
    inc(ARichIndex);
  end;

  AComplete := (ARichIndex >= nRtp);
end;


{@@ ----------------------------------------------------------------------------
  Write the result of a string formula in the preceding record.
  In BIFF8 files no STRING record occurs, if the result string is empty.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteSTRINGRecord(AStream: TStream;
  AString: String);
// wp: This method might be imcomplete:
// - Missing call to FixLineEnding()
// - Missing RichText
// - Missing check for length (max 8224 bytes per record, else use CONTINUE)
var
  wideStr: widestring;
  len: Integer;
  strBytes: Integer;
  idx: Integer;
  needCONTINUE: Boolean;
begin
  if AString = '' then
    exit;

  wideStr := WideStringToLE(UTF8Decode(FixLineEnding(AString)));
  len := Length(wideStr);

  strBytes := len * SizeOf(WideChar);
  needCONTINUE := 3 + strBytes > MAX_BYTES_IN_RECORD;

  if needCONTINUE then
    strBytes := MAX_BYTES_IN_RECORD - 4;  // -4 = -3 (header) - 1 (even byte count)

  { BIFF STRING record header}
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_STRING));
  AStream.WriteWord(WordToLE(3 + strBytes));

  { Write widestring length }
  AStream.WriteWord(WordToLE(len));
  { Widestring flags, 1=regular unicode LE string }
  AStream.WriteByte(1);
  { Write characters }
  AStream.WriteBuffer(wideStr[1], strBytes);

  idx := 1 + strBytes div SizeOf(WideChar);

  while needCONTINUE and (idx < len) do begin
    strBytes := (len - idx) * SizeOf(WideChar);
    needCONTINUE := strBytes + 1 > MAX_BYTES_IN_RECORD;
    if needCONTINUE then
      strBytes := MAX_BYTES_IN_RECORD - 2;  // -2 = -1 (flag byte) - 1 (for even count)

    { BIFF CONTINUE record header }
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_CONTINUE));
    AStream.WriteWord(WordToLE(1 + strBytes)); // for flag byte
    { Widestring flags, 1 = regular unicode LE string }
    AStream.WriteByte(1);
    {Write characters }
    AStream.WriteBuffer(wideStr[idx], strBytes);

    inc(idx, strBytes div SizeOf(WideChar));
  end;
end;

{@@-----------------------------------------------------------------------------
  Writes an Excel 8 STYLE record

  Registers the name of a user-defined style or specific options
  for a built-in cell style.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteSTYLE(AStream: TStream);
begin
  { BIFF record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_STYLE, 4);

  { Index to style XF and defines if it's a built-in or used defined style }
  AStream.WriteWord(WordToLE(MASK_STYLE_BUILT_IN));

  { Built-in cell style identifier }
  AStream.WriteByte($00);

  { Level if the identifier for a built-in style is RowLevel or ColLevel, $FF otherwise }
  AStream.WriteByte($FF);
end;

{@@ ----------------------------------------------------------------------------
  Writes a TXO and two CONTINUE records as needed for cell comments.
  It can safely be assumed that the cell exists and contains a comment.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteTXO(AStream: TStream; AComment: PsComment);
var
  recTXO: TBIFF8TXORecord;
  comment: widestring;
  compressed: ansistring;
  len: Integer;
  wchar: widechar;
  i: Integer;
  bytesFmtRuns: Integer;
begin
  { Prepare comment string. It is stored as a string with 8-bit characters }
  comment := UTF8Decode(AComment^.Text);
  SetLength(compressed, length(comment));
  for i:= 1 to Length(comment) do
  begin
    wchar := comment[i];
    compressed[i] := wchar;
  end;
  len := Length(compressed);

  { (1)  TXO record ---------------------------------------------------------- }
  { BIFF record header }
  FillChar(recTXO{%H-}, SizeOf(recTXO), 0);
  recTXO.RecordID := WordToLE(INT_EXCEL_ID_TXO);
  recTXO.RecordSize := SizeOf(recTXO) - 2*SizeOf(word);
  { Record data }
  recTXO.OptionFlags := WordToLE($0212);  // Left & top aligned, lock option on
  recTXO.TextRot := 0;                    // Comment text not rotated
  recTXO.TextLen := WordToLE(len);
  bytesFmtRuns := 8*SizeOf(Word);         // see (3) below
  recTXO.NumFormattingRuns := WordToLE(bytesFmtRuns);
  { Write out to file }
  AStream.WriteBuffer(recTXO, SizeOf(recTXO));

  { (2)  1st CONTINUE record containing the comment text --------------------- }
  { BIFF record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_CONTINUE));
  AStream.WriteWord(len+1);
  { Record data }
  AStream.WriteByte(0);
  AStream.WriteBuffer(compressed[1], len);

  { (3)  2nd CONTINUE record containing the formatting runs ------------------ }
  { BIFF record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_CONTINUE));
  AStream.WriteWord(bytesFmtRuns);
  { Record data }
  AStream.WriteWord(0);             // start index of 1st formatting run (we only use 1 run)
  AStream.WriteWord(WordToLE(1));   // Font index to be used (default font)
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(WordToLE(len)); // lastRun: number of characters
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(0);             // Not used
  AStream.WriteWord(0);             // Not used
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 WINDOW2 record

  This record contains additional settings for the document window (BIFF2-BIFF4)
  or for a specific worksheet (BIFF5-BIFF8).

  The values written here are reasonable defaults, which should work for most
  sheets.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteWINDOW2(AStream: TStream;
  ASheet: TsWorksheet);
var
  Options: Word;
  actSheet: TsWorksheet;
begin
  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_WINDOW2, 18);

  { Options flags }
  Options :=
    MASK_WINDOW2_OPTION_SHOW_ZERO_VALUES or
    MASK_WINDOW2_OPTION_AUTO_GRIDLINE_COLOR or
    MASK_WINDOW2_OPTION_SHOW_OUTLINE_SYMBOLS;
{or
    MASK_WINDOW2_OPTION_SHEET_SELECTED or
    MASK_WINDOW2_OPTION_SHEET_ACTIVE;}
   { Bug 0026386 -> every sheet must be selected/active, otherwise Excel cannot print
     ---> wp: after changes for issue 0028452: this is not necessary any more. }

  if (soShowGridLines in ASheet.Options) then
    Options := Options or MASK_WINDOW2_OPTION_SHOW_GRID_LINES;
  if (soShowHeaders in ASheet.Options) then
    Options := Options or MASK_WINDOW2_OPTION_SHOW_SHEET_HEADERS;
  if (soHasFrozenPanes in ASheet.Options) and ((ASheet.LeftPaneWidth > 0) or (ASheet.TopPaneHeight > 0)) then
    Options := Options or MASK_WINDOW2_OPTION_PANES_ARE_FROZEN;
  if FWorkbook.ActiveWorksheet <> nil then
    actSheet := FWorkbook.ActiveWorksheet else
    actSheet := Fworkbook.GetWorksheetByIndex(0);
  if (ASheet = actSheet) then
    Options := Options or MASK_WINDOW2_OPTION_SHEET_ACTIVE or MASK_WINDOW2_OPTION_SHEET_SELECTED;
  if (ASheet.BiDiMode = bdRTL) then
    Options := Options or MASK_WINDOW2_OPTION_COLUMNS_RIGHT_TO_LEFT;
  AStream.WriteWord(WordToLE(Options));

  { Index to first visible row }
  AStream.WriteWord(WordToLE(0));

  { Index to first visible column }
  AStream.WriteWord(WordToLE(0));

  { Grid line index colour }
  AStream.WriteWord(WordToLE(0));

  { Not used }
  AStream.WriteWord(WordToLE(0));

  { Cached magnification factor in page break preview (in percent); 0 = Default (60%) }
  AStream.WriteWord(WordToLE(0));

  { Cached magnification factor in normal view (in percent); 0 = Default (100%) }
  AStream.WriteWord(WordToLE(0));

  { Not used }
  AStream.WriteDWord(DWordToLE(0));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 8 XF record (cell format)
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF8Writer.WriteXF(AStream: TStream;
  AFormatRecord: PsCellFormat; XFType_Prot: Byte = 0);
var
  rec: TBIFF8_XFRecord;
  j: Integer;
  b: Byte;
  dw1, dw2: DWord;
  w3: Word;
  nfParams: TsNumFormatParams;
  nfs: String;
begin
  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_XF);
  rec.RecordSize := WordToLE(SizeOf(TBIFF8_XFRecord) - SizeOf(TsBIFFHeader));

  { Index to font record }
  rec.FontIndex := 0;
  if (AFormatRecord <> nil) then begin
    if (uffFont in AFormatRecord^.UsedFormattingFields) then
    begin
      rec.FontIndex := AFormatRecord^.FontIndex;
      if rec.FontIndex >= 4 then inc(rec.FontIndex);  // Font #4 does not exist in BIFF
    end;
  end;
  rec.FontIndex := WordToLE(rec.FontIndex);

  { Index to number format }
  j := 0;
  if (AFormatRecord <> nil) and (uffNumberFormat in AFormatRecord^.UsedFormattingFields)
  then begin
    nfParams := Workbook.GetNumberFormat(AFormatRecord^.NumberFormatIndex);
    if nfParams <> nil then
    begin
      nfs := nfParams.NumFormatStr;
      j := NumFormatList.IndexOf(nfs);
      if j = -1 then j := 0;
    end;
  end;
  rec.NumFormatIndex := WordToLE(j);

  { XF type, cell protection and parent style XF }
  if AFormatRecord = nil then begin
    rec.XFType_Prot_ParentXF := XFType_Prot and MASK_XF_TYPE_PROT;
    if XFType_Prot and MASK_XF_TYPE_PROT_STYLE_XF <> 0 then
      rec.XFType_Prot_ParentXF := rec.XFType_Prot_ParentXF or MASK_XF_TYPE_PROT_PARENT;
  end else
  begin
    rec.XFType_Prot_ParentXF := 0;
    if cpLockCell in AFormatRecord^.Protection then
      rec.XFType_Prot_ParentXF := rec.XFType_Prot_ParentXF or MASK_XF_TYPE_PROT_LOCKED;
    if cpHideFormulas in AFormatRecord^.Protection then
      rec.XFType_Prot_ParentXF := rec.XFType_Prot_ParentXF or MASK_XF_TYPE_PROT_FORMULA_HIDDEN;
  end;
  rec.XFType_Prot_ParentXF := WordToLE(rec.XFType_Prot_ParentXF);

  { Text alignment and text break }
  if AFormatRecord = nil then
    b := MASK_XF_VERT_ALIGN_BOTTOM
  else
  begin
    b := 0;
    if (uffHorAlign in AFormatRecord^.UsedFormattingFields) then
      case AFormatRecord^.HorAlignment of
        haDefault: ;
        haLeft   : b := b or MASK_XF_HOR_ALIGN_LEFT;
        haCenter : b := b or MASK_XF_HOR_ALIGN_CENTER;
        haRight  : b := b or MASK_XF_HOR_ALIGN_RIGHT;
      end;
    // Since the default vertical alignment is vaDefault but "0" corresponds
    // to vaTop, we alwys have to write the vertical alignment.
    case AFormatRecord^.VertAlignment of
      vaTop    : b := b or MASK_XF_VERT_ALIGN_TOP;
      vaCenter : b := b or MASK_XF_VERT_ALIGN_CENTER;
      vaBottom : b := b or MASK_XF_VERT_ALIGN_BOTTOM;
      else       b := b or MASK_XF_VERT_ALIGN_BOTTOM;
    end;
    if (uffWordWrap in AFormatRecord^.UsedFormattingFields) then
      b := b or MASK_XF_TEXTWRAP;
  end;
  rec.Align_TextBreak := b;

  { Text rotation }
  rec.TextRotation := 0;
  if (AFormatRecord <> nil) and (uffTextRotation in AFormatRecord^.UsedFormattingFields)
    then rec.TextRotation := TEXT_ROTATIONS[AFormatRecord^.TextRotation];

  { Indentation, shrink, merge and text direction:
    see "Excel97-2007BinaryFileFormat(xls)Specification.pdf", p281 ff
    Bits 0-3: Indent value
    Bit 4: Shrink to fit
    Bit 5: MergeCell
    Bits 6-7: Reading direction  }
  rec.Indent_Shrink_TextDir := 0;
  if (AFormatRecord <> nil) and (uffBiDi in AFormatRecord^.UsedFormattingFields) then
  begin
    b := ord(AFormatRecord^.BiDiMode);
    if b > 0 then
      rec.Indent_Shrink_TextDir := rec.Indent_Shrink_TextDir or (b shl 6);
  end;

  { Used attributes }
  rec.UsedAttrib :=
    MASK_XF_USED_ATTRIB_NUMBER_FORMAT or
    MASK_XF_USED_ATTRIB_FONT or
    MASK_XF_USED_ATTRIB_TEXT or
    MASK_XF_USED_ATTRIB_BORDER_LINES or
    MASK_XF_USED_ATTRIB_BACKGROUND or
    MASK_XF_USED_ATTRIB_CELL_PROTECTION;

  { Cell border lines and background area }

  dw1 := 0;
  dw2 := 0;
  w3 := 0;
  if (AFormatRecord <> nil) and (uffBorder in AFormatRecord^.UsedFormattingFields) then
  begin
    // Left and right line colors
    dw1 := PaletteIndex(AFormatRecord^.BorderStyles[cbWest].Color) shl 16 +
           PaletteIndex(AFormatRecord^.BorderStyles[cbEast].Color) shl 23;
    // Border line styles
    if cbWest in AFormatRecord^.Border then
      dw1 := dw1 or (DWord(AFormatRecord^.BorderStyles[cbWest].LineStyle)+1);
    if cbEast in AFormatRecord^.Border then
      dw1 := dw1 or ((DWord(AFormatRecord^.BorderStyles[cbEast].LineStyle)+1) shl 4);
    if cbNorth in AFormatRecord^.Border then
      dw1 := dw1 or ((DWord(AFormatRecord^.BorderStyles[cbNorth].LineStyle)+1) shl 8);
    if cbSouth in AFormatRecord^.Border then
      dw1 := dw1 or ((DWord(AFormatRecord^.BorderStyles[cbSouth].LineStyle)+1) shl 12);
    if cbDiagDown in AFormatRecord^.Border then
      dw1 := dw1 or $40000000;
    if cbDiagUp in AFormatRecord^.Border then
      dw1 := dw1 or $80000000;

    // Top, bottom and diagonal line colors
    dw2 := PaletteIndex(AFormatRecord^.BorderStyles[cbNorth].Color) +
           PaletteIndex(AFormatRecord^.BorderStyles[cbSouth].Color) shl 7 +
           PaletteIndex(AFormatRecord^.BorderStyles[cbDiagUp].Color) shl 14;
    // In BIFF8 both diagonals have the same color - we use the color of the up-diagonal.

    // Diagonal line style
    if (AFormatRecord^.Border * [cbDiagUp, cbDiagDown] <> []) then
      dw2 := dw2 or ((DWord(AFormatRecord^.BorderStyles[cbDiagUp].LineStyle)+1) shl 21);
    // In BIFF8 both diagonals have the same line style - we use the color of the up-diagonal.
  end;

  { Background fill }
  if (AFormatRecord <> nil) and (uffBackground in AFormatRecord^.UsedFormattingFields) then
  begin
    // Fill pattern style
    dw2 := dw2 or DWORD(MASK_XF_FILL_PATT[AFormatRecord^.Background.Style] shl 26);
    // Pattern color
    if AFormatRecord^.Background.FgColor = scTransparent
      then w3 := w3 or SYS_DEFAULT_FOREGROUND_COLOR
      else w3 := w3 or PaletteIndex(AFormatRecord^.Background.FgColor);
    // Background color
    if AFormatRecord^.Background.BgColor = scTransparent
      then w3 := w3 or SYS_DEFAULT_BACKGROUND_COLOR shl 7
      else w3 := w3 or (PaletteIndex(AFormatRecord^.Background.BgColor) shl 7);
  end;

  rec.Border_BkGr1 := DWordToLE(dw1);
  rec.Border_BkGr2 := DWordToLE(dw2);
  rec.BkGr3 := WordToLE(w3);

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;


initialization

  // Registers this reader / writer in fpSpreadsheet
  sfidExcel8 := RegisterSpreadFormat(sfExcel8,
    TsSpreadBIFF8Reader, TsSpreadBIFF8Writer,
    STR_FILEFORMAT_EXCEL_8, 'BIFF8', [STR_EXCEL_EXTENSION]
  );

  // Converts the palette to litte-endian
  MakeLEPalette(PALETTE_BIFF8);

end.

