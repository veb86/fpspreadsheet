unit fpsImages;

{$mode objfpc}{$H+}

interface

uses
  Classes;

type
  TGetImageSizeFunc = function (AStream: TStream;
    out AWidth, AHeight: DWord; out dpiX, dpiY: Double): Boolean;

  TsImageType = integer;

const
  {@@ Identifier for unknown image type }
  itUnknown = -1;

var
  {@@ Identifier for the PNG image type (value 0) }
  itPNG: TsImageType;
  {@@ Identifier for the JPEG image type (value 1) }
  itJPEG: TsImageType;
  {@@ Identifier for the TIFF image type (value 2) }
  itTIFF: TsImageType;
  {@@ Identifier for the BMP image type (value 3) }
  itBMP: TsImageType;
  {@@ Identifier for the GIF image type (value 4) }
  itGIF: TsImageType;
  {@@ Identifier for the SVG image type (value 5) }
  itSVG: TsImageType;
  {@@ Identifier for the WMF image type (value 6) }
  itWMF: TsImageType;
  {@@ Identifier for the EMF image type (value 7) }
  itEMF: TsImageType;
  {@@ Identifier for the PCX image type (value 8) }
  itPCX: TsImageType;

type
  { TsEmbeddedObj }
  TsEmbeddedObj = class
  private
    FFileName: String;
    FStream: TMemoryStream;
    FImageType: TsImageType;  // image type, see itXXXX
    FWidth: Double;           // image width, in mm
    FHeight: Double;          // image height, in mm
  protected
    function CheckStream(AImageType: TsImageType): Boolean;
  public
    destructor Destroy; override;
    function LoadFromFile(const AFileName: String): Boolean;
    function LoadFromStream(AStream: TStream; AName: String;
      ASize: Int64 = -1): Boolean;
    property FileName: String read FFileName;
    property ImageType: TsImagetype read FImageType;
    property ImageWidth: Double read FWidth write FWidth;
    property ImageHeight: Double read FHeight write FHeight;
    property Stream: TMemoryStream read FStream;
  end;


function GetImageInfo(AStream: TStream; out AWidthInches, AHeightInches: Double;
  AImagetype: TsImageType = itUnknown): TsImageType; overload;

function GetImageInfo(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double; AImageType: TsImageType = itUnknown): TsImageType; overload;

function GetImageMimeType(AImageType: TsImageType): String;
function GetImageTypeExt(AImageType: TsImageType): String;
function GetImageTypeFromFileName(const AFilename: String): TsImageType;

function RegisterImageType(AMimeType, AExt: String; AGetImageSize: TGetImageSizeFunc): TsImageType;


implementation

uses
  SysUtils, Strings, Math,
  fpsUtils;

type
  TByteOrder = (boLE, boBE);  // little edian, or big endian

  TImageTypeRecord = record
    Ext: String;
    MimeType: String;
    GetImageSize: TGetImageSizeFunc;
  end;

var
  ImageTypeRegistry: array of TImageTypeRecord;

{ Makes sure that the byte order of w is as specified by the parameter }
function FixByteOrder(w: Word; AByteOrder: TByteOrder): Word; overload;
begin
  Result := IfThen(AByteOrder = boLE, LEToN(w), BEToN(w));
end;

{ Makes sure that the byte order of dw is as specified by the parameter }
function FixByteOrder(dw: DWord; AByteOrder: TByteOrder): DWord; overload;
begin
  Result := IfThen(AByteOrder = boLE, LEToN(dw), BEToN(dw));
end;

function GetTIFSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean; forward;


{ BMP files }

function GetBMPSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
// stackoverflow.com/questions/15209076/how-to-get-dimensions-of-image-file-in-delphi
type
  TBitMapFileHeader = packed record
     bfType: word;
     bfSize: longint;
     bfReserved: longint;
     bfOffset: longint;
  end;
  TBitMapInfoHeader = packed record
     Size: longint;
     Width: longint;
     Height: longint;
     Planes: word;
     BitCount: word;
     Compression: longint;
     SizeImage: longint;
     XPelsPerMeter: Longint;
     YPelsPerMeter: Longint;
     ClrUsed: longint;
     ClrImportant: longint;
  end;
const
  BMP_MAGIC_WORD = ord('M') shl 8 or ord('B');
var
  header: TBitmapFileHeader;
  info: TBitmapInfoHeader;
begin
  result := False;
  dpiX := 0;
  dpiY := 0;
  if AStream.Read(header{%H-}, SizeOf(header)) <> SizeOf(header) then Exit;
  if LEToN(header.bfType) <> BMP_MAGIC_WORD then Exit;
  if AStream.Read(info{%H-}, SizeOf(info)) <> SizeOf(info) then Exit;
  AWidth := LEToN(info.Width);
  AHeight := abs(LEToN(info.Height));  // can be negative in case of "top-down" image
  if info.Size >= 40 then
  begin
    dpiX := LEToN(info.XPelsPerMeter) * 0.0254;
    dpiY := LEToN(info.YPelsPerMeter) * 0.0254;
  end;
  if dpiX = 0 then dpiX := 72;
  if dpiY = 0 then dpiY := 72;
  Result := true;
end;


{ EMF files }

function GetEMFSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
// https://msdn.microsoft.com/de-de/library/windows/desktop/dd162607%28v=vs.85%29.aspx
type
  TEnhMetaHeader = packed record
     iType: DWord;
     nSize: DWord;
     rclBounds: TRect;
     rclFrame: TRect;
     dSignature: DWord;   // must be $464D4520
     nVersion: DWord;
     nBytes: DWord;
     nRecords: DWord;
     nHandles: Word;
     sReserved: Word;
     nDescription: DWord;
     offDescription: DWord;
     nPalEntries: DWord;
     szlDevice: TPoint;
     szlMillimeters: TPoint;
     // more to follow
  end;
var
  hdr: TEnhMetaHeader;
  n: Int64;
begin
  Result := false;

  n := AStream.Read(hdr{%H-}, SizeOf(hdr));
  if n < SizeOf(hdr) then exit;
  if hdr.dSignature <> $464D4520 then exit;

  AWidth := (hdr.rclFrame.Right - hdr.rclFrame.Left);  // in 0.01 mm
  AHeight := (hdr.rclFrame.Bottom - hdr.rclFrame.Top);
  dpiX := 100*25.4;
  dpiY := 100*25.4;

  Result := true;
end;


{ GIF files }

function GetGIFSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
type
  TGifHeader = packed record
    Sig: array[0..5] of char;
    ScreenWidth, ScreenHeight: word;
    Flags, Background, Aspect: byte;
  end;
  TGifImageBlock = packed record
    Left, Top, Width, Height: word;
    Flags: byte;
  end;
var
  header: TGifHeader;
  imageBlock: TGifImageBlock;
  nResult: integer;
  x: integer;
  c: char;
begin
  Result := false;

  // Read header and ensure valid file
  nResult := AStream.Read(header{%H-}, SizeOf(TGifHeader));
  if (nResult <> SizeOf(TGifHeader)) then exit;  // invalid file
  if (strlicomp(PChar(header.Sig), 'GIF87a', 6) <> 0) and
     (strlicomp(PChar(header.Sig), 'GIF89a', 6) <> 0) then exit;

  // Skip color map, if there is one
  if (header.Flags and $80) > 0 then
  begin
    x := 3 * (1 SHL ((header.Flags and 7) + 1));
    AStream.Position := x;
    if AStream.Position > AStream.Size then exit; // Color map thrashed
  end;
  // Step through blocks
  while (AStream.Position < AStream.Size) do
  begin
    c := char(AStream.ReadByte);
    if c = ',' then
    begin
      // Image found
      nResult := AStream.Read(imageBlock{%H-}, SizeOf(TGIFImageBlock));
      if nResult <> SizeOf(TGIFImageBlock) then exit; // Invalid image block encountered
      AWidth := LEToN(imageBlock.Width);
      AHeight := LEToN(imageBlock.Height);
      break;
    end;
  end;
  dpiX := 96;   // not stored in file, use default screen dpi
  dpiY := 96;
  Result := true;
end;


{ JPG files }

function GetJPGSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): boolean;
type
  TJPGHeader = array[0..1] of Byte; //FFD8 = StartOfImage (SOI)
  TJPGRecord = packed record
    Marker: Byte;
    RecType: Byte;
    RecSize: Word;
  end;
  TAPP0Record = packed record
    JFIF: Array[0..4] of AnsiChar;  // zero-terminated "JFIF" string
    Version: Word;     // JFIF format revision
    Units: Byte;       // Units used for resolution: 1->inch, 2->cm, 0-> aspect ratio (1, 1)
    XDensity: Word;    // Horizontal resolution
    YDensity: Word;    // Vertical resolution
    // thumbnail follows
  end;
var
  n: integer;
  hdr: TJPGHeader;
  rec: TJPGRecord = (Marker: $FF; RecType: 0; RecSize: 0);
  app0: TAPP0Record;
  u: Integer;
  p: Int64;
  exifSig: Array[0..5] of AnsiChar;
  imgW, imgH: DWord;
begin
  Result := false;

  AWidth := 0;
  AHeight := 0;
  dpiX := -1;
  dpiY := -1;
  u := -1;  // units of pixel density

  // Check for SOI (start of image) record
  n := AStream.Read(hdr{%H-}, SizeOf(hdr));
  if (n < SizeOf(hdr)) or (hdr[0] <> $FF) or (hdr[1] <> $D8) then
    exit;

  while (AStream.Position < AStream.Size) and (rec.Marker = $FF) do begin
    if AStream.Read(rec, SizeOf(rec)) < SizeOf(rec) then exit;
    rec.RecSize := BEToN(rec.RecSize);
    p := AStream.Position - 2;
    case rec.RecType of
      $E0:  // APP0 record
        if (rec.RecSize >= SizeOf(TAPP0Record)) then
        begin
          AStream.Read(app0{%H-}, SizeOf(app0));
          if stricomp(pchar(app0.JFIF), 'JFIF') <> 0 then break;
          dpiX := BEToN(app0.XDensity);
          dpiY := BEToN(app0.YDensity);
          u := app0.Units;
        end else
          exit;
      $E1:   // APP1 record (EXIF)
        begin
          AStream.Read(exifSig{%H-}, Sizeof(exifSig));
          if not GetTIFSize(AStream, imgW, imgH, dpiX, dpiY) then exit;
        end;
      $C0..$C3:
        if (rec.RecSize >= 4) then // Start of frame markers
        begin
          AStream.Seek(1, soFromCurrent);  // Skip "bits per sample"
          AHeight := BEToN(AStream.ReadWord);
          AWidth := BEToN(AStream.ReadWord);
        end else
          exit;
      $D9:  // end of image;
        break;
    end;
    AStream.Position := p + rec.RecSize;
  end;

  if (dpiX = -1) or (u = 0) then dpiX := 96;
  if (dpiY = -1) or (u = 0) then dpiY := 96;
  if u = 2 then begin
    dpiX := dpiX * 2.54;
    dpiY := dpiY * 2.54;
  end;

  Result := true;
end;


{ PCX files }

function GetPCXSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
type
  TPCXHeader = packed record
    FileID: Byte;                      // $0A for PCX files, $CD for SCR files
    Version: Byte;                     // 0: version 2.5; 2: 2.8 with palette; 3: 2.8 w/o palette; 5: version 3
    Encoding: Byte;                    // 0: uncompressed; 1: RLE encoded
    BitsPerPixel: Byte;
    XMin,
    YMin,
    XMax,
    YMax,                              // coordinates of the corners of the image
    HRes,                              // horizontal resolution in dpi
    VRes: Word;                        // vertical resolution in dpi
    ColorMap: array[0..15*3] of byte;  // color table
    Reserved,
    ColorPlanes: Byte;                 // color planes (at most 4)
    BytesPerLine,                      // number of bytes of one line of one plane
    PaletteType: Word;                 // 1: color or b&w; 2: gray scale
    Fill: array[0..57] of Byte;
  end;
var
  hdr: TPCXHeader;
  n: Int64;
begin
  Result := false;

  n := AStream.Read(hdr{%H-}, SizeOf(hdr));
  if n < SizeOf(hdr) then exit;
  if not (hdr.FileID in [$0A, $CD]) then exit;

  AWidth := hdr.XMax - hdr.XMin + 1;
  AHeight := hdr.YMax - hdr.YMin + 1;
  dpiX := hdr.HRes;
  dpiY := hdr.VRes;

  Result := True;
end;


{ PNG files }

function GetPNGSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
// https://www.w3.org/TR/PNG/
type
  TPngSig = array[0..7] of byte;
  TPngChunk = packed record
    chLength: LongInt;
    chType: array[0..3] of AnsiChar;
  end;
const
  ValidSig: TPNGSig = (137, 80, 78, 71, 13, 10, 26, 10);
var
  Sig: TPNGSig;
  x: integer;
  chunk: TPngChunk;
  xdpm: LongInt;
  ydpm: LongInt;
  units: Byte;
  p: Int64;
begin
  Result := false;
  dpiX := 96;
  dpiY := 96;

  FillChar(Sig{%H-}, SizeOf(Sig), #0);
  AStream.Read(Sig[0], SizeOf(Sig));
  for x := Low(Sig) to High(Sig) do
    if Sig[x] <> ValidSig[x] then
      exit;
  AStream.Seek(18, 0);
  AWidth := BEToN(AStream.ReadWord);
  AStream.Seek(22, 0);
  AHeight := BEToN(AStream.ReadWord);
  AStream.Position := SizeOf(TPngSig);
  while AStream.Position < AStream.Size do
  begin
    AStream.Read(chunk{%H-}, SizeOf(TPngChunk));
    chunk.chLength := BEToN(chunk.chLength);
    p := AStream.Position;
    if strlcomp(PChar(chunk.chType), 'pHYs', 4) = 0 then
    begin
      xdpm := BEToN(AStream.ReadDWord);  // pixels per meter
      ydpm := BEToN(AStream.ReadDWord);
      units := AStream.ReadByte;
      if units = 1 then
      begin
        dpiX := xdpm * 0.0254;
        dpiY := ydpm * 0.0254;
      end;
      break;
    end;
    AStream.Position := p + chunk.chLength + 4;
  end;
  Result := true;
end;


{ SVG files }

function GetSVGSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
var
  fs: TFormatSettings;

  function Extract(AName, AText: String): String;
  var
    p: Integer;
  begin
    Result := '';
    p := pos(lowercase(AName), lowercase(AText));
    if p > 0 then
    begin
      inc(p, Length(AName));
      while (p <= Length(AText)) and (AText[p] in [' ', '"', '=']) do
        inc(p);
      while (p <= Length(AText)) and (AText[p] <> '"') do
      begin
        Result := Result + AText[p];
        inc(p);
      end;
    end;
  end;

  function ToInches(AText: String): Double;
  begin
    if AText[Length(AText)] in ['0'..'9'] then
      Result := mmToIn(StrToFloat(AText, fs))
    else
      Result := PtsToIn(HTMLLengthStrToPts(AText));
  end;

  // Split the 4 viewbox values. If values don't have attached units assume mm.
  // Return viewbox width and height in inches.
  function AnalyzeViewbox(AText: String; out w, h: Double): Boolean;
  var
    L: TStringList;
  begin
    L := TStringList.Create;
    try
      L.Delimiter := ' ';
      L.StrictDelimiter := true;
      L.DelimitedText := AText;
      if L.Count <> 4 then exit(false);

      w := ToInches(L[2]) - ToInches(L[0]);
      h := ToInches(L[3]) - ToInches(L[1]);

      Result := true;
    finally
      L.Free;
    end;
  end;

var
  ch: AnsiChar;
  s: String;
  done: Boolean;
  sW, sH, sVB: String;
  w, h: Double;
begin
  Result := false;
  AWidth := 0;
  AHeight := 0;

  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';

  // Assume 100 dpi --> Multiply the inches by 100
  dpiX := 100;
  dpiY := 100;

  done := false;
  while (not done) and (AStream.Position < AStream.Size) do
  begin
    ch := char(AStream.ReadByte);
    if ch = '<' then begin
      ch := char(AStream.ReadByte);
      if ch <> 's' then continue;
      ch := char(AStream.ReadByte);
      if ch <> 'v' then continue;
      ch := char(AStream.ReadByte);
      if ch <> 'g' then continue;
      ch := char(AStream.ReadByte);
      if ch <> ' ' then continue;
      s := '';
      while (not done) and (AStream.Position < AStream.Size) do
      begin
        ch := char(AStream.Readbyte);
        if ch = '>' then
          done := true
        else
          s := s + ch;
      end;
    end;
  end;
  if not done then
    exit;

  sW := Extract('width', s);
  sH := Extract('height', s);
  sVB := Extract('viewBox', s);

  // If "viewBox" exists, ignore "Width" and "Height" except for percentage
  if (sVB <> '') and AnalyzeViewBox(sVB, w, h) then
  begin
    if (sW <> '') and (sW[Length(sw)] = '%') then begin
      SetLength(sW, Length(sW)-1);
      AWidth := round(w * StrToFloat(sW, fs) / 100 * dpiX);
    end else
      AWidth := round(w * dpiX);
    if (sH <> '') and (sH[Length(sH)] = '%') then begin
      SetLength(sH, Length(sH)-1);
      AHeight := round(h * StrToFloat(sH, fs) / 100 * dpiY);
    end else
      AHeight := round(h * dpiY);
  end else
  begin
    if sw <> '' then
      AWidth := round(HTMLLengthStrToPts(sW) * 72 * dpiX);
    if sh <> '' then
      AHeight := round(HTMLLengthStrToPts(sH) * 72 * dpiY);
  end;

  Result := true;
end;


{ TIF files }

function GetTIFSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
type
  TTifHeader = packed record
     BOM: word;     // 'II' for little endian, 'MM' for big endian
     Sig: word;     // Signature (42)
     IFD: DWORD;    // Offset where image data begin
  end;
  TIFD_Field = packed record
    Tag: word;
    FieldType: word;
    ValCount: DWord;
    ValOffset: DWord;
  end;
var
  header: TTifHeader = (BOM:0; Sig:0; IFD:0);
  dirEntries: Word;
  field: TIFD_Field = (Tag:0; FieldType:0; ValCount:0; ValOffset:0);
  i: Integer;
  bo: TByteOrder;
  num, denom: LongInt;
  units: Word;
  p, pStart: Int64;
begin
  Result := false;
  AWidth := 0;
  AHeight := 0;
  dpiX := 0;
  dpiY := 0;
  units := 0;

  // Remember current stream position because procedure is called also from
  // jpeg Exif block.
  pStart := AStream.Position;

  if AStream.Read(header, SizeOf(TTifHeader)) < SizeOf(TTifHeader) then exit;
  if not ((header.BOM = $4949) or (header.BOM = $4D4D)) then exit;
  if header.BOM = $4949 then bo := boLE else bo := boBE; // 'II' --> little endian, 'MM' --> big endian
  if FixByteOrder(header.Sig, bo) <> 42 then exit;

  AStream.Position := pStart + FixByteOrder(header.IFD, bo);
  dirEntries := FixByteOrder(AStream.ReadWord, bo);
  for i := 1 to dirEntries do
  begin
    AStream.Read(field, SizeOf(field));
    field.Tag := FixByteOrder(field.Tag, bo);
    field.ValOffset := FixByteOrder(field.ValOffset, bo);
    field.FieldType := FixByteOrder(field.FieldType, bo);
    p := AStream.Position;
    case field.Tag OF
      $0100 : AWidth := field.ValOffset;
      $0101 : AHeight := field.ValOffset;
      $011A : begin    // XResolution as RATIONAL value
                AStream.Position := pStart + field.ValOffset;
                num := FixByteOrder(AStream.ReadDWord, bo);
                denom := FixByteOrder(AStream.ReadDWord, bo);
                dpiX := num/denom;
              end;
      $011B : begin    // YResolution as RATIONAL value
                AStream.Position := pStart + field.ValOffset;
                num := FixByteOrder(AStream.ReadDWord, bo);
                denom := FixByteOrder(AStream.ReadDWord, bo);
                dpiY := num/denom;
              end;
      $0128 : begin
                units := field.ValOffset;   // 1: non-square 2: inches, 3: cm
              end;
    end;
    if (AWidth > 0) and (AHeight > 0) and (dpiX > 0) and (dpiY > 0) and (units > 0)
    then
      break;
    AStream.Position := p;
  end;

  case units of
    1: begin dpiX := 96; dpiY := 96; end;
    2: ;  // is already inches, nothing to do
    3: begin dpiX := dpiX*2.54; dpiY := dpiY * 2.54; end;
  end;

  Result := true;
end;


{ WMF files }

function GetWMFSize(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double): Boolean;
type
  TWMFSpecialHeader = packed record
    Key: DWord;       // Magic number (always $9AC6CDD7)
    Handle: Word;     // Metafile HANDLE number (always 0)
    Left: SmallInt;   // Left coordinate in metafile units (twips)
    Top: SmallInt;    // Top coordinate in metafile units
    Right: SmallInt;  // Right coordinate in metafile units
    Bottom: SmallInt; // Bottom coordinate in metafile units
    Inch: Word;       // Number of metafile units per inch
    Reserved: DWord;  // Reserved (always 0)
    Checksum: Word;   // Checksum value for previous 10 words
  end;
var
  hdr: TWMFSpecialHeader;
  n: Int64;
begin
  Result := false;

  n := AStream.Read(hdr{%H-}, SizeOf(hdr));
  if n < SizeOf(hdr) then exit;
  if hdr.Key <> $9AC6CDD7 then exit;

  AWidth := (hdr.Right - hdr.Left);
  AHeight := (hdr.Bottom - hdr.Top);
  dpiX := hdr.Inch;
  dpiY := hdr.Inch;

  Result := true;
end;


{==============================================================================}
{                           Public functions                                   }
{==============================================================================}

{@@ ----------------------------------------------------------------------------
  Returns the width and height of the image loaded into the specified stream.

  @param  AStream       Stream containing the image to be analyzed. It is
                        assumed that the image begins at stream start.
  @param  AWidthInches  Image width, in inches
  @param  AHeightInches Image height, in inches
  @param  AImageType    Type of the image to be assumed. If this parameter is
                        missing or itUnknown then the image type is determined
                        from the file header.

  @return Image type code found from the file header.
  @see RegisterImageType
-------------------------------------------------------------------------------}
function GetImageInfo(AStream: TStream; out AWidthInches, AHeightInches: Double;
  AImagetype: TsImageType = itUnknown): TsImageType;
var
  w, h: DWord;
  xdpi, ydpi: Double;
begin
  Result := GetImageInfo(AStream, w, h, xdpi, ydpi, AImageType);
  if Result <> itUnknown then begin
    AWidthInches := w / xdpi;
    AHeightInches := h / ydpi;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the width and height of the image loaded into the specified stream.

  @param  AStream       Stream containing the image to be analyzed. It is
                        assumed that the image begins at stream start.
  @param  AWidth        Image width, in pixels
  @param  AHeight       Image height, in pixels
  @param  dpiX          Pixel density in x direction, per inch
  @param  dpiY          Pixel density in y direction, per inch
  @param  AImageType    Type of the image to be assumed. If this parameter is
                        missing or itUnknown then the image type is determined
                        from the file header.

  @return Image type code found from the file header.
  @see RegisterImageType
-------------------------------------------------------------------------------}
function GetImageInfo(AStream: TStream; out AWidth, AHeight: DWord;
  out dpiX, dpiY: Double; AImageType: TsImageType = itUnknown): TsImageType;
var
  itr: TImageTypeRecord;  // [i]mage [t]ype [r]ecord
begin
  if InRange(AImageType, 0, High(ImageTypeRegistry)) then
  begin
    AStream.Position := 0;
    if ImageTypeRegistry[AImageType].GetImageSize(AStream, AWidth, AHeight, dpiX, dpiY)
      then Result := AImageType;
  end else
  begin
    for Result := 0 to High(ImageTypeRegistry) do
    begin
      AStream.Position := 0;
      itr := ImageTypeRegistry[Result];
      if itr.GetImageSize(AStream, AWidth, AHeight, dpiX, dpiY) then
        exit;
    end;
    Result := itUnknown;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the MimeType of the specified image type

  @param   AImageType    Format code of the image type as returned from the
                         image registration procedure
  @return  MimeType of the file format
-------------------------------------------------------------------------------}
function GetImageMimeType(AImageType: TsImageType): String;
begin
  if InRange(AImageType, 0, High(ImageTypeRegistry)) then
    Result := ImageTypeRegistry[AImageType].MimeType
  else
    Result := '';
end;

{@@ ----------------------------------------------------------------------------
  Returns the file extension belonging the specified image type. If there
  are several extensions the first one is selected. The extension is returned
  without a leading period.
-------------------------------------------------------------------------------}
function GetImageTypeExt(AImageType: TsImageType): String;
var
  p: Integer;
begin
  if InRange(AImageType, 0, High(ImageTypeRegistry)) then
  begin
    Result := ImageTypeRegistry[AImageType].Ext;
    p := pos('|', Result);
    if p > 0 then
      Result := copy(Result, 1, p-1);
    if Result[1] = '.' then Delete(Result, 1, 1);
  end else
    Result := '';
end;

{@@ ----------------------------------------------------------------------------
  Extracts the image file type identifier from the extension of the specified
  file name.

  @param  AFileName    Name of the file to be analyzed
  @return Format code value as returned from the image registration procedure
  @see    RegisterImageType, itXXXX values.
-------------------------------------------------------------------------------}
function GetImageTypeFromFileName(const AFilename: String): TsImageType;
var
  ext: String;
  i,j: Integer;
  itr: TImageTypeRecord;
  regext: TStringArray;
begin
  ext := Lowercase(ExtractFileExt(AFilename));
  Delete(ext, 1, 1);
  for j := 0 to High(ImageTypeRegistry) do
  begin
    itr := ImageTypeRegistry[j];
    regext := SplitStr(itr.Ext, '|');
    for i := 0 to High(regext) do
      if regext[i] = ext then
      begin
        Result := TsImageType(j);
        exit;
      end;
  end;
  Result := itUnknown;
end;

{@@ ----------------------------------------------------------------------------
  Registers an image type for usage in fpspreadsheet

  @param  AExt            Extension(s) of the file format. Separate by "|" if a
                          file format can use several extensions.
  @param  AMimeType       MimeType of the file format, for usage by ods
  @param  AGetImageSize   Function which can extract the image size and
                          pixel density. It should only read the file header.
  @return  Identifier of the image type (consecutive number)
-------------------------------------------------------------------------------}
function RegisterImageType(AMimeType, AExt: String;
  AGetImageSize: TGetImageSizeFunc): TsImageType;
begin
  Result := Length(ImageTypeRegistry);
  SetLength(ImageTypeRegistry, Result + 1);
  with ImageTypeRegistry[Result] do
  begin
    MimeType := AMimeType;
    Ext := AExt;
    GetImageSize := AGetImageSize;
  end;
end;


{==============================================================================}
{                               TsEmbeddedObj                                  }
{==============================================================================}

destructor TsEmbeddedObj.Destroy;
begin
  FreeAndNil(FStream);
  inherited Destroy;
end;

function TsEmbeddedObj.CheckStream(AImageType: TsImageType): Boolean;
begin
  FImageType := GetImageInfo(FStream, FWidth, FHeight, AImageType);
  // FWidth and FHeight are in inches here.
  Result := FImageType <> itUnknown;
end;

function TsEmbeddedObj.LoadFromFile(const AFileName: String): Boolean;
var
  s: TStream;
begin
  FreeAndNil(FStream);
  FStream := TMemoryStream.Create;
  s := TFileStream.Create(AFileName, fmOpenRead + fmShareDenyNone);
  try
    FStream.LoadFromStream(s);
    Result := CheckStream(GetImageTypeFromFileName(AFileName));
    if Result then FFileName := AFileName;
  finally
    s.Free;
  end;
end;

function TsEmbeddedObj.LoadFromStream(AStream: TStream; AName: String;
  ASize: Int64 = -1): Boolean;
begin
  FreeAndNil(FStream);
  FStream := TMemoryStream.Create;
  if ASize = -1 then
    FStream.LoadFromStream(AStream)
  else
    FStream.CopyFrom(AStream, ASize);
  Result := CheckStream(itUnknown);
  if Result then FFileName := AName;
end;



initialization

{0}  itPNG  := RegisterImageType('image/png',  'png', @GetPNGSize);
{1}  itJPEG := RegisterImageType('image/jpeg', 'jpg|jpeg|jfif|jfe', @GetJPGSize);
{2}  itTIFF := RegisterImageType('image/tiff', 'tif|tiff', @GetTIFSize);
{3}  itBMP  := RegisterImageType('image/bmp', 'bmp|dib', @GetBMPSize);
{4}  itGIF  := RegisterImageType('image/gif', 'gif', @GetGIFSize);
{5}  itSVG  := RegisterImageType('image/svg+xml', 'svg', @GetSVGSize);
{6}  itWMF  := RegisterImageType('application/x-msmetafile', 'wmf', @GetWMFSize);
{7}  itEMF  := RegisterImageType('image/x-emf', 'emf', @GetEMFSize);
{8}  itPCX  := RegisterImageType('image/pcx', 'pcx', @GetPCXSize);

end.
