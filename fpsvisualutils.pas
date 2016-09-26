unit fpsvisualutils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics,
  fpstypes, fpspreadsheet;

procedure Convert_sFont_to_Font(sFont: TsFont; AFont: TFont); overload;
procedure Convert_sFont_to_Font(AWorkbook: TsWorkbook; sFont: TsFont; AFont: TFont); overload; deprecated;

procedure Convert_Font_to_sFont(AFont: TFont; sFont: TsFont); overload;
procedure Convert_Font_to_sFont(AWorkbook: TsWorkbook; AFont: TFont; sFont: TsFont); overload; deprecated;

function WrapText(ACanvas: TCanvas; const AText: string; AMaxWidth: integer): string;

procedure DrawRichText(ACanvas: TCanvas; AWorkbook: TsWorkbook; const ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  AWordwrap: Boolean; AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation; AOverrideTextColor: TColor; ARightToLeft: Boolean;
  AZoomFactor: Double);

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean;
  AZoomFactor: Double): Integer;

function RichTextHeight(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean;
  AZoomFactor: Double): Integer;

type
  TsLineInfo = class
    pStart: PChar;
    WordList: TStringList;
    NumSpaces: Integer;
    BeginsWithFontOfRtpIndex: Integer;
    Width: Integer;
    Height: Integer;
    constructor Create;
    destructor Destroy; override;
  end;

  { TsTextPainter }

  TsTextPainter = class
  private
    FCanvas: TCanvas;
    FWorkbook: TsWorkbook;
    FRect: TRect;
    FFontIndex: Integer;
    FTextRotation: TsTextRotation;
    FHorAlignment: TsHorAlignment;
    FVertAlignment: TsVertAlignment;
    FWordWrap: Boolean;
    FRightToLeft: Boolean;
    FText: String;
    FRtParams: TsRichTextParams;
    FMaxLineLen: Integer;
    FTotalHeight: Integer;
    FLines: TFPList;
    FPtr: PChar;
    FRtpIndex: Integer;
    FCharIndex: integer;
    FCharIndexOfNextFont: Integer;
    FFontHeight: Integer;
    FFontPos: TsFontPosition;
    FZoomFactor: Double;

  private
    function GetHeight: Integer;
    function GetWidth: Integer;

  protected
    procedure DrawHor(AOverrideTextColor: TColor);
    procedure DrawLine(pEnd: PChar; x, y, ALineHeight: Integer; AOverrideTextColor: TColor);
    procedure DrawStacked(AOverrideTextColor: TColor);
    procedure DrawText(var x, y: Integer; s: String; ALineHeight: Integer);
    procedure DrawVert(AOverrideTextColor: TColor; AClockwise: Boolean);

    function GetTextPt(x,y,ALineHeight: Integer): TPoint;
    procedure InitFont(out ACurrRtpIndex, ACharIndexOfNextFont, ACurrFontHeight: Integer;
      out ACurrFontPos: TsFontPosition);
    procedure NextChar(ANumBytes: Integer);
    procedure Prepare;
    procedure ScanLine(var ANumSpaces, ALineWidth, ALineHeight: Integer;
      AWordList: TStringList);
    procedure UpdateFont(ACharIndex: Integer; var ACurrRtpIndex,
      ACharIndexOfNextFont, ACurrFontHeight: Integer; var ACurrFontPos: TsFontPosition);

  public
    constructor Create(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
      AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
      ATextRotation: TsTextRotation; AHorAlignment: TsHorAlignment;
      AVertAlignment: TsVertAlignment; AWordWrap, ARightToLeft: Boolean;
      AZoomFactor: Double);
    destructor Destroy; override;
    procedure Draw(AOverrideTextColor: TColor);
    property Height: Integer read GetHeight;
    property Width: Integer read GetWidth;
  end;

implementation

uses
  Types, Math, LCLType, LCLIntf, LazUTF8, fpsUtils;

const
{@@ Font size factor for sub-/superscript characters }
  SUBSCRIPT_SUPERSCRIPT_FACTOR = 0.66;

{@@ ----------------------------------------------------------------------------
  Converts a spreadsheet font to a font used for painting (TCanvas.Font).

  @param  sFont      Font as used by fpspreadsheet (input)
  @param  AFont      Font as used by TCanvas for painting (output)
-------------------------------------------------------------------------------}
procedure Convert_sFont_to_Font(sFont: TsFont; AFont: TFont);
begin
  if Assigned(AFont) and Assigned(sFont) then begin
    AFont.Name := sFont.FontName;
    AFont.Size := round(sFont.Size);
    AFont.Style := [];
    if fssBold in sFont.Style then AFont.Style := AFont.Style + [fsBold];
    if fssItalic in sFont.Style then AFont.Style := AFont.Style + [fsItalic];
    if fssUnderline in sFont.Style then AFont.Style := AFont.Style + [fsUnderline];
    if fssStrikeout in sFont.Style then AFont.Style := AFont.Style + [fsStrikeout];
    AFont.Color := TColor(sFont.Color and $00FFFFFF);
  end;
end;

procedure Convert_sFont_to_Font(AWorkbook: TsWorkbook; sFont: TsFont; AFont: TFont);
begin
  Unused(AWorkbook);
  Convert_sFont_to_Font(sFont, AFont);
end;

{@@ ----------------------------------------------------------------------------
  Converts a font used for painting (TCanvas.Font) to a spreadsheet font.

  @param  AFont  Font as used by TCanvas for painting (input)
  @param  sFont  Font as used by fpspreadsheet (output)
-------------------------------------------------------------------------------}
procedure Convert_Font_to_sFont(AFont: TFont; sFont: TsFont);
begin
  if Assigned(AFont) and Assigned(sFont) then begin
    sFont.FontName := AFont.Name;
    sFont.Size := AFont.Size;
    sFont.Style := [];
    if fsBold in AFont.Style then Include(sFont.Style, fssBold);
    if fsItalic in AFont.Style then Include(sFont.Style, fssItalic);
    if fsUnderline in AFont.Style then Include(sFont.Style, fssUnderline);
    if fsStrikeout in AFont.Style then Include(sFont.Style, fssStrikeout);
    sFont.Color := ColorToRGB(AFont.Color);
  end;
end;

procedure Convert_Font_to_sFont(AWorkbook: TsWorkbook; AFont: TFont; sFont: TsFont);
begin
  Unused(AWorkbook);
  Convert_Font_to_sFont(AFont, sFont);
end;

{@@ ----------------------------------------------------------------------------
  Wraps text by inserting line ending characters so that the lines are not
  longer than AMaxWidth.

  @param   ACanvas       Canvas on which the text will be drawn
  @param   AText         Text to be drawn
  @param   AMaxWidth     Maximimum line width (in pixels)
  @return  Text with inserted line endings such that the lines are shorter than
           AMaxWidth.

  @note    Based on ocde posted by user "taazz" in the Lazarus forum
           http://forum.lazarus.freepascal.org/index.php/topic,21305.msg124743.html#msg124743
-------------------------------------------------------------------------------}
function WrapText(ACanvas: TCanvas; const AText: string; AMaxWidth: integer): string;
var
  DC: HDC;
  textExtent: TSize = (cx:0; cy:0);
  S, P, E: PChar;
  line: string;
  isFirstLine: boolean;
begin
  Result := '';
  DC := ACanvas.Handle;
  isFirstLine := True;
  P := PChar(AText);
  while P^ = ' ' do
    Inc(P);
  while P^ <> #0 do begin
    S := P;
    E := nil;
    while (P^ <> #0) and (P^ <> #13) and (P^ <> #10) do begin
      LCLIntf.GetTextExtentPoint(DC, S, P - S + 1, textExtent);
      if (textExtent.CX > AMaxWidth) and (E <> nil) then begin
        if (P^ <> ' ') and (P^ <> ^I) then begin
          while (E >= S) do
            case E^ of
              '.', ',', ';', '?', '!', '-', ':',
              ')', ']', '}', '>', '/', '\', ' ':
                break;
              else
                Dec(E);
            end;
          if E < S then
            E := P - 1;
        end;
        Break;
      end;
      E := P;
      Inc(P);
    end;
    if E <> nil then begin
      while (E >= S) and (E^ = ' ') do
        Dec(E);
    end;
    if E <> nil then
      SetString(Line, S, E - S + 1)
    else
      SetLength(Line, 0);
    if (P^ = #13) or (P^ = #10) then begin
      Inc(P);
      if (P^ <> (P - 1)^) and ((P^ = #13) or (P^ = #10)) then
        Inc(P);
      if P^ = #0 then
        line := line + LineEnding;
    end
    else if P^ <> ' ' then
      P := E + 1;
    while P^ = ' ' do
      Inc(P);
    if isFirstLine then begin
      Result := Line;
      isFirstLine := False;
    end else
      Result := Result + LineEnding + line;
  end;
end;


{------------------------------------------------------------------------------}
{                       Public rich-text functios                              }
{------------------------------------------------------------------------------}

procedure DrawRichText(ACanvas: TCanvas; AWorkbook: TsWorkbook; const ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  AWordwrap: Boolean; AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation; AOverrideTextColor: TColor; ARightToLeft: Boolean;
  AZoomFactor: Double);
var
  painter: TsTextPainter;
begin
  if (ARect.Left = ARect.Right) or (ARect.Top = ARect.Bottom) then
    exit;

  painter := TsTextPainter.Create(ACanvas, AWorkbook, ARect, AText, ARichTextParams,
    AFontIndex, ARotation, AHorAlignment, AVertAlignment, AWordWrap, ARightToLeft,
    AZoomFactor);
  try
    painter.Draw(AOverrideTextColor);
  finally
    painter.Free;
  end;
end;

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean;
  AZoomFactor: Double): Integer;
var
  painter: TsTextPainter;
begin
  if (ARect.Left = ARect.Right) or (ARect.Top = ARect.Bottom) then
    exit(0);

  painter := TsTextPainter.Create(ACanvas, AWorkbook, ARect, AText, ARichTextParams,
    AFontIndex, ATextRotation, haLeft, vaTop, AWordWrap, ARightToLeft, AZoomFactor);
  try
    Result := painter.Width;
  finally
    painter.Free;
  end;
end;

function RichTextHeight(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean;
  AZoomFactor: Double): Integer;
var
  painter: TsTextPainter;
begin
  if (ARect.Left = ARect.Right) or (ARect.Top = ARect.Bottom) then
    exit(0);

  painter := TsTextPainter.Create(ACanvas, AWorkbook, ARect, AText, ARichTextParams,
    AFontIndex, ATextRotation, haLeft, vaTop, AWordWrap, ARightToLeft, AZoomFactor);
  try
    Result := painter.Height;
  finally
    painter.Free;
  end;
end;


{------------------------------------------------------------------------------}
{                    Painting engine for rich-text                             }
{------------------------------------------------------------------------------}

constructor TsLineInfo.Create;
begin
  inherited;
  WordList := TStringList.Create;
end;

destructor TsLineInfo.Destroy;
begin
  WordList.Free;
  inherited;
end;


{ TsTextPainter }

{ ARect ........ Defines the rectangle in which the text is to be drawn,
  AFontIndex ... Base font of the text, to be used if not rich-text is defined.
  ATextRoation . Text is rotated this way
  AWordwrap .... Wrap text at word boundaries if text is wider than the MaxRect
                 (or higher, in case of vertical text).
  ARightToLeft . if true, paint text from left to right }
constructor TsTextPainter.Create(ACanvas: TCanvas; AWorkbook: TsWorkbook;
  ARect: TRect; AText: String; ARichTextParams: TsRichTextParams;
  AFontIndex: Integer; ATextRotation: TsTextRotation; AHorAlignment: TsHorAlignment;
  AVertAlignment: TsVertAlignment; AWordWrap, ARightToLeft: Boolean;
  AZoomFactor: Double);
begin
  FLines := TFPList.Create;
  FCanvas := ACanvas;
  FWorkbook := AWorkbook;
  FRect := ARect;
  FText := AText;
  FRtParams := ARichTextParams;
  FFontIndex := AFontIndex;
  FTextRotation := ATextRotation;
  FHorAlignment := AHorAlignment;
  FVertAlignment := AVertAlignment;
  FWordwrap := AWordwrap;
  FRightToLeft := ARightToLeft;
  FZoomfactor := AZoomFactor;
  Prepare;
end;

destructor TsTextPainter.Destroy;
var
  j: Integer;
begin
  for j := FLines.Count-1 downto 0 do TObject(FLines[j]).Free;
  FLines.Free;
  inherited Destroy;
end;

{ Draw the lines }
procedure TsTextPainter.Draw(AOverrideTextColor: TColor);
begin
  case FTextRotation of
    trHorizontal                       : DrawHor(AOverrideTextColor);
    rt90DegreeClockwiseRotation        : DrawVert(AOverrideTextColor, true);
    rt90DegreeCounterClockwiseRotation : DrawVert(AOverrideTextColor, false);
    rtStacked                          : DrawStacked(AOverrideTextColor);
  end;
end;

{ Draw lines in horizontal orienation }
procedure TsTextPainter.DrawHor(AOverrideTextColor: TColor);
var
  xpos, ypos, j: Integer;
  lineinfo: TsLineInfo;
  pEnd: PChar;
begin
  // (1) Get starting point of line
  case FVertAlignment of
    vaTop    : ypos := FRect.Top;
    vaCenter : ypos := (FRect.Top + FRect.Bottom - FTotalHeight) div 2;
    vaBottom : ypos := FRect.Bottom - FTotalHeight;
  end;

  // (2) Draw text line-by-line
  FPtr := PChar(FText);
  FCharIndex := 1;
  InitFont(FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
  for j := 0 to FLines.Count-1 do
  begin
    if j < FLines.Count-1 then
      pEnd := TsLineInfo(FLines[j+1]).pStart else
      pEnd := PChar(FText) + Length(FText);
    lineinfo := TsLineInfo(FLines[j]);
    // xpos is x coordinate of left edge of first character
    if FRightToLeft then
      case FHorAlignment of
        haLeft   : xpos := FRect.Left + lineinfo.Width;
        haCenter : xpos := (FRect.Left + FRect.Right + lineinfo.Width) div 2;
        haRight  : xpos := FRect.Right;
      end
    else
      case FHorAlignment of
        haLeft   : xpos := FRect.Left;
        haCenter : xpos := (FRect.Left + FRect.Right - lineinfo.Width) div 2;
        haRight  : xpos := FRect.Right - lineinfo.Width;
      end;
    DrawLine(pEnd, xpos, ypos, lineinfo.Height, AOverrideTextColor);
    inc(ypos, lineinfo.Height);
  end;
end;

{ Draw a single line. The font can change within the line. }
procedure TsTextPainter.DrawLine(pEnd: PChar; x, y, ALineHeight: Integer;
  AOverrideTextColor: TColor);
var
  charLen: Integer;
  s: String;
begin
  s := '';
  while (FPtr^ <> #0) and (FPtr < pEnd) do begin
    if FCharIndex = FCharIndexOfNextFont then begin
      DrawText(x, y, s, ALineHeight);
      s := '';
    end;
    UpdateFont(FCharIndex, FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
    if AOverrideTextColor <> clNone then
      FCanvas.Font.Color := AOverrideTextColor;
    case FPtr^ of
      #10: begin
             DrawText(x, y, s, ALineHeight);
             s := '';
             NextChar(1);
             break;
           end;
      #13: begin
             DrawText(x, y, s, ALineHeight);
             s := '';
             NextChar(1);
             if FPtr^ = #10 then
               NextChar(1);
             break;
           end;
      else
           s := s + UnicodeToUTF8(UTF8CharacterToUnicode(FPtr, charLen));
           if FCharIndex = FCharIndexOfNextFont then begin
             DrawText(x, y, s, ALineHeight);
             s := '';
           end;
           NextChar(charLen);
    end;
  end;
  if s <> '' then
    DrawText(x, y, s, ALineHeight);
end;

// Draws text in vertical columns using upright characters
procedure TsTextPainter.DrawStacked(AOverrideTextColor: TColor);
const
  IGNORE = 0;
var
  xpos, ypos, dx: Integer;
  j: Integer;
  lineinfo: TsLineInfo;
  pEnd: PChar;
begin
  // (1) Get starting point of line
  lineinfo := TsLineInfo(FLines[0]);
  dx := lineInfo.Height;
  if FRightToLeft then
    case FHorAlignment of
      haLeft   : xpos := FRect.Left + FTotalHeight + dx;
      haCenter : xpos := (FRect.Left + FRect.Right + FTotalHeight) div 2 - dx;
      haRight  : xpos := FRect.Right - dx;
    end
  else
    case FHorAlignment of
      haLeft   : xpos := FRect.Left + dx;
      haCenter : xpos := (FRect.Left + FRect.Right - FTotalHeight) div 2;
      haRight  : xpos := FRect.Right - FTotalHeight + dx;
    end;

  // (2) Draw text line-by-line
  FPtr := PChar(FText);
  FCharIndex := 1;
  InitFont(FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
  for j := 0 to FLines.Count-1 do
  begin
    if j < FLines.Count-1 then
      pEnd := TsLineInfo(FLines[j+1]).pStart
    else
      pEnd := PChar(FText) + Length(FText);
    lineinfo := TsLineInfo(FLines[j]);
    case FVertAlignment of
      vaTop    : ypos := FRect.Top;
      vaCenter : ypos := (FRect.Top + FRect.Bottom - lineinfo.Width) div 2;
      vaBottom : ypos := FRect.Bottom - lineinfo.Width;
    end;
    DrawLine(pEnd, xpos, ypos, IGNORE, AOverrideTextColor);
    if FRightToLeft then
      dec(xpos, 2*lineinfo.Height) else        // "height" is horizontal here!
      inc(xpos, 2*lineinfo.Height);
  end;
end;

{ Draw a text chunk. Font does not change here }
procedure TsTextPainter.DrawText(var x, y: Integer; s: String;
  ALineHeight: Integer);
const
  MULTIPLIER: Array[TsTextRotation, boolean] of Integer = (
    (+1, -1),  // horiz                ^
    (+1, -1),  // 90° CW           FRightToLeft
    (-1, +1),  // 90° CCW
    (+1, -1)   // stacked
  );
  TEXT_ANGLE: array[TsTextRotation] of Integer = ( 0, -900, 900, 0);
var
  w, wlead, wtrail: Integer;
  Pt: TPoint;
  i, nlead, ntrail, nchar: Integer;
  p: PChar;
  charLen: Integer;
  ch: String;
begin
  wlead := 0;
  wtrail := 0;
  if FRightToLeft then
  begin
    { Right-to-left character handling of RTL strings containing spaces is very
      confusing -- probably this is not correct... }
    // Count leading spaces
    nlead := 0;
    i := 1;
    while (i <= Length(s)) and (s[i] = ' ') do begin
      inc(i);
      inc(nlead);
    end;
    wlead := nlead * FCanvas.TextWidth(' ');
    // count trailing spaces
    ntrail := 0;
    i := Length(s);
    while (i >= 1) and (s[i] = ' ') do begin
      dec(i);
      inc(ntrail);
    end;
    wtrail := ntrail * FCanvas.TextWidth(' ');
    // Remove leading and trailing spaces from string; their size will be
    // compensated by coordinate offset wlead/wtrail.
    s := trim(s);
  end;
  w := FCanvas.TextWidth(s);
  Pt := GetTextPt(x, y, ALineHeight);
  FCanvas.Font.Orientation := TEXT_ANGLE[FTextRotation];
  case FTextRotation of
    trHorizontal:
      begin
        if FRightToLeft
          then FCanvas.TextOut(Pt.x-w-wlead, Pt.y, s)
          else FCanvas.TextOut(Pt.x, Pt.y, s);
        inc(x, (wlead+w+wtrail)*MULTIPLIER[FTextRotation, FRightToLeft]);
      end;
    rt90DegreeClockwiseRotation:
      begin
        if FRightToLeft
          then FCanvas.TextOut(Pt.x, Pt.y-w-wlead, s)
          else FCanvas.TextOut(Pt.x, Pt.y, s);
        inc(y, (wlead+w+wtrail)*MULTIPLIER[FTextRotation, FRightToLeft]);
      end;
    rt90DegreeCounterClockwiseRotation:
      begin
        if FRightToLeft
          then FCanvas.TextOut(Pt.x, Pt.y+w+wlead, s)
          else FCanvas.TextOut(Pt.x, Pt.y, s);
        inc(y, (wlead+w+wtrail)*MULTIPLIER[FTextRotation, FRightToLeft]);
      end;
    rtStacked:
      begin
        nChar := 0;
        P := PChar(s);
        while (P^ <> #0) do
        begin
          ch := UnicodeToUTF8(UTF8CharacterToUnicode(P, charLen));
          ALineHeight := FCanvas.TextHeight(ch);
          Pt := GetTextPt(x, y, ALineHeight);
          w := FCanvas.TextWidth(ch);
          // x is at the center of the character here
          case FHorAlignment of
            haLeft   : FCanvas.TextOut(Pt.x - w div 2, Pt.y, ch);
            haCenter : FCanvas.TextOut(Pt.x - w div 2, Pt.y, ch);
            haRight  : FCanvas.TextOut(Pt.x - w div 2, Pt.y, ch);
          end;
          inc(y, ALineHeight);
          inc(nChar);
          inc(P, charLen);
        end;
      end;
  end;
end;

// Draw text in 90° clockwise or counter-clockwise rotation
procedure TsTextPainter.DrawVert(AOverrideTextColor: TColor; AClockwise: Boolean);
const                           // CCW  CW
  SGN: array[boolean] of Integer = (-1, +1);
var
  j, xpos, ypos: Integer;
  lineinfo: TsLineInfo;
  pEnd: PChar;
begin
  // (1) Get starting point
  case FHorAlignment of
    haLeft   : xpos := IfThen(AClockwise, FRect.Left + FTotalHeight, FRect.Left);
    haCenter : xpos := (FRect.Left + FRect.Right + FTotalHeight*SGN[AClockwise]) div 2;
    haRight  : xpos := IfThen(AClockwise, FRect.Right, FRect.Right - FTotalHeight);
  end;

  // (2) Draw text line by line and respect text rotation
  FPtr := PChar(FText);
  FCharIndex := 1;      // Counter for utf8 character position
  InitFont(FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
  for j := 0 to FLines.Count-1 do
  begin
    if j < FLines.Count-1 then
      pEnd := TsLineInfo(FLines[j+1]).pStart else
      pEnd := PChar(FText) + Length(FText);
    lineinfo := TsLineInfo(FLines[j]);
    if FRightToLeft then
      case FVertAlignment of
        vaTop    : ypos := IfThen(AClockwise, FRect.Top + lineinfo.Width, FRect.Top);
        vaCenter : ypos := (FRect.Top + FRect.Bottom + lineinfo.Width*SGN[AClockwise]) div 2;
        vaBottom : ypos := IfThen(AClockwise, FRect.Bottom, FRect.Bottom - lineinfo.Width);
      end
    else
      case FVertAlignment of
        vaTop    : ypos := IfThen(AClockwise, FRect.Top, FRect.Top + lineinfo.Width);
        vaCenter : ypos := (FRect.Top + FRect.Bottom - lineinfo.Width*SGN[AClockwise]) div 2;
        vaBottom : ypos := IfThen(AClockwise, FRect.Bottom - lineinfo.Width, FRect.Bottom);
      end;
    DrawLine(pEnd, xpos, ypos, lineinfo.Height, AOverrideTextColor);
    inc(xpos, -lineinfo.Height*SGN[AClockwise]);
  end;
end;

function TsTextPainter.GetHeight: Integer;
begin
  if FTextRotation = trHorizontal then
    Result := FTotalHeight
  else
    Result := FMaxLineLen;
end;

function TsTextPainter.GetTextPt(x,y,ALineHeight: Integer): TPoint;
begin
  case FTextRotation of
    trHorizontal, rtStacked:
      case FFontPos of
        fpNormal      : Result := Point(x, y);
        fpSubscript   : Result := Point(x, y + ALineHeight div 2);
        fpSuperscript : Result := Point(x, y - ALineHeight div 6);
      end;
    rt90DegreeClockwiseRotation:
      case FFontPos of
        fpNormal      : Result := Point(x, y);
        fpSubscript   : Result := Point(x - ALineHeight div 2, y);
        fpSuperscript : Result := Point(x + ALineHeight div 6, y);
      end;
    rt90DegreeCounterClockWiseRotation:
      case FFontPos of
        fpNormal      : Result := Point(x, y);
        fpSubscript   : Result := Point(x + ALineHeight div 2, y);
        fpSuperscript : Result := Point(x - ALineHeight div 6, y);
      end;
  end;
end;

function TsTextPainter.GetWidth: Integer;
begin
  if FTextRotation = trHorizontal then
    Result := FMaxLineLen else
    Result := FTotalHeight;
end;

{ Called before analyzing and rendering of the text.
  ACurrRtpIndex ......... Index of CURRENT rich-text parameter
  ACharIndexOfNextFont .. Character index when NEXT font change will occur
  ACurrFontHeight ....... CURRENT font height
  ACurrFontPos .......... CURRENT font position (normal/sub/superscript) }
procedure TsTextPainter.InitFont(out ACurrRtpIndex, ACharIndexOfNextFont,
  ACurrFontHeight: Integer; out ACurrFontPos: TsFontPosition);
var
  fnt: TsFont;
begin
  FCharIndex := 1;
  if (Length(FRtParams) = 0) then
  begin
    FRtpIndex := -1;
    fnt := FWorkbook.GetFont(FFontIndex);
    ACharIndexOfNextFont := MaxInt;
  end
  else if (FRtParams[0].FirstIndex = 1) then
  begin
    ACurrRtpIndex := 0;
    fnt := FWorkbook.GetFont(FRtParams[0].FontIndex);
    if Length(FRtParams) > 1 then
      ACharIndexOfNextFont := FRtParams[1].FirstIndex
    else
      ACharIndexOfNextFont := MaxInt;
  end else
  begin
    fnt := FWorkbook.GetFont(FFontIndex);
    ACurrRtpIndex := -1;
    ACharIndexOfNextFont := FRtParams[0].FirstIndex;
  end;
  Convert_sFont_to_Font(fnt, FCanvas.Font);
  FCanvas.Font.Height := round(FZoomFactor * FCanvas.Font.Height);
  ACurrFontHeight := FCanvas.TextHeight('Tg');
  if (fnt <> nil) and (fnt.Position <> fpNormal) then
    FCanvas.Font.Size := round(fnt.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR * FZoomFactor);
  ACurrFontPos := fnt.Position;
end;

procedure TsTextPainter.NextChar(ANumBytes: Integer);
begin
  inc(FPtr, ANumBytes);
  inc(FCharIndex);
end;

{ Get layout of lines
  "lineinfos" collect data for where lines start and end, their width and
  height, the rich-text parameter index range, and the number of spaces and
  a word list (for text justification). }
procedure TsTextPainter.Prepare;
var
  lineInfo: TsLineInfo;
  ts: TTextStyle;
begin
  FTotalHeight := 0;
  FMaxLinelen := 0;

  if FText = '' then
    exit;

  ts := FCanvas.TextStyle;
  ts.RightToLeft := FRightToLeft;
  FCanvas.TextStyle := ts;

  InitFont(FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);

  FPtr := PChar(FText);
  FCharIndex := 1;
  while (FPtr^ <> #0) do begin
    lineInfo := TsLineInfo.Create;
    lineInfo.pStart := FPtr;
    lineInfo.BeginsWithFontOfRtpIndex := FRtpIndex;
    ScanLine(lineInfo.NumSpaces, lineInfo.Width, lineInfo.Height, lineInfo.WordList);
    FLines.Add(lineinfo);
    FTotalHeight := FTotalHeight + IfThen(FTextRotation = rtStacked, 2, 1)*lineInfo.Height;
    FMaxLineLen := Max(FMaxLineLen, lineInfo.Width);
  end;
end;

{ Scans the line for a possible line break and a font change.
  The scan starts at the current position of FPtr.

  ANumSpaces     is how many spaces were found between the start and end value
                 of FPtr.
  ALineWidth     the pixel width of the line seen along drawing direction, i.e.
                 in case of stacked text it is the sum of the character heights!
  ALineHeight    The height of the line as seen vertically to the drawing
                 direction. Normally this is the height of the largest font
                 found in the line; in case of stacked text it is the
                 width of character 'M'. }
procedure TsTextPainter.ScanLine(var ANumSpaces, ALineWidth, ALineHeight: Integer;
  AWordList: TStringList);
var
  tmpWidth: Integer;
  savedWidth: Integer;
  savedHeight: Integer;
  savedSpaces: Integer;
  savedCharIndex: Integer;
  savedCurrRtpIndex: Integer;
  savedCharIndexOfNextFont: Integer;
  maxWidth: Integer;
  s: String;
  charLen: Integer;
  ch: String;
  EOL: Boolean;
  pWordStart: PChar;
  part, savedPart: String;
begin
  ANumSpaces := 0;
  ALineHeight := FFontHeight;
  ALineWidth := 0;
  savedWidth := 0;
  savedHeight := 0;
  savedSpaces := 0;
  s := '';      // current word
  part := '';   // current part of the string where all characters have the same font
  savedpart := '';
  tmpWidth := 0;

  maxWidth := MaxInt;
  if FWordWrap then
  begin
    if FTextRotation = trHorizontal then
      maxWidth := FRect.Right - FRect.Left
    else
      maxWidth := FRect.Bottom - FRect.Top;
  end;

  while (FPtr^ <> #0) do
  begin
    case FPtr^ of
      #13: begin
             NextChar(1);
             if FPtr^ = #10 then
               NextChar(1);
             break;
           end;
      #10: begin
             NextChar(1);
             break;
           end;
      ' ': begin
             ALineWidth := ALineWidth + tmpWidth;
             part := '';
             tmpWidth := 0;  // width of the spaces, growing during scan
             // Save data for the case that max width is exceeded here
             savedWidth := ALineWidth;
             savedHeight := ALineHeight;
             savedSpaces := ANumSpaces;
             savedPart := part;
             // Find next word
             while FPtr^ = ' ' do
             begin
               // We reached a character at which the font changes
               // --> update current line width
               // This has to be done before "UpdateFont" because the collected
               // part string uses the old font.
               if (FCharIndex = FCharIndexOfNextFont) then
               begin
                 if (FTextRotation <> rtStacked) then
                   tmpwidth := tmpwidth + FCanvas.TextWidth(part);
                 part := '';
                 savedPart := '';
                 tmpwidth := 0;
               end;
               // Update font if required
               UpdateFont(FCharIndex, FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
               part := part + ' ';
               if FTextRotation = rtStacked then
               begin
                 tmpwidth := tmpwidth + FFontHeight;
                 ALineHeight := Max(ALineHeight, FCanvas.TextWidth('M'));
               end else begin
                 tmpwidth := tmpwidth + FCanvas.TextWidth(' ');
                 ALineHeight := Max(ALineHeight, FFontHeight);
               end;
               inc(ANumSpaces);
               NextChar(1);
             end;
             if ALineWidth + tmpWidth <= maxWidth then
             begin
               if FTextRotation = rtStacked then
                 ALineWidth := ALineWidth + tmpWidth;
             end else
             begin
               // max width has been exceeded while scanning spaces
               // --> restore values stored at the end of previous word
               ALineWidth := savedWidth;
               ALineHeight := savedHeight;
               ANumSpaces := savedSpaces;
               part := savedPart;
               while (part <> '') and (part[Length(part)] = ' ') do
               begin
                 Delete(part, Length(part), 1);
                 if FTextRotation = rtStacked then dec(ALineWidth, FFontHeight);
               end;
               break;
             end;
           end;
      else
           // Here, a new word begins. Find the end of this word and check if
           // it fits into the line.
           // Store the data valid for the word start. They are needed if the
           // scan would go beyond the max line width in this word.
           s := '';
           pWordStart := FPtr;
           savedCharIndex := FCharIndex;
           savedCurrRtpIndex := FRtpIndex;
           savedCharIndexOfNextFont := FCharIndexOfNextFont;
           savedpart := part;
           savedHeight := ALineHeight;
           tmpWidth := 0;  // width of the current word, growing during the scan
           EOL := false;
           while (FPtr^ <> #0) and (FPtr^ <> #13) and (FPtr^ <> #10) and (FPtr^ <> ' ') do
           begin
             if FCharIndex = FCharIndexOfNextFont then
             begin
               if (FTextRotation <> rtStacked) then
                 ALineWidth := ALineWidth + FCanvas.TextWidth(part);
               part := '';
               tmpWidth := 0;
             end;
             UpdateFont(FCharIndex, FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
             ch := UnicodeToUTF8(UTF8CharacterToUnicode(FPtr, charLen));
             part := part + ch;
             if FTextRotation = rtStacked then
             begin
               tmpWidth := tmpWidth + FFontHeight;
               ALineHeight := Max(ALineHeight, FCanvas.TextWidth('M'));
             end else
             begin
               tmpWidth := FCanvas.TextWidth(part);
               ALineHeight := Max(FFontHeight, ALineHeight);
             end;
             if ALineWidth + tmpWidth <= maxWidth then
               s := s + ch
             else
             begin
               // The line exeeds the max line width.
               // There are two cases:
               if ANumSpaces > 0 then
               begin
                 // (a) This is not the only word: Go back to where this
                 // word began. We already had stored everything needed!
                 FPtr := pWordStart;
                 FCharIndex := savedCharIndex;
                 FCharIndexOfNextFont := savedCharIndexOfNextFont;
                 FRtpIndex := savedCurrRtpIndex;
                 ALineHeight := savedHeight;
                 part := savedPart;
                 while (part <> '') and (part[Length(part)] = ' ') do
                 begin
                   Delete(part, Length(part), 1);
                   if FTextRotation = rtStacked then dec(ALineWidth, FFontHeight);
                 end;
               end else
               begin
                 // (b) This is the only word in the line --> we break at the
                 // current cursor position.
                 if Length(part) = 1 then
                   NextChar(1)
                 else
                   UTF8Delete(part, UTF8Length(part), 1);
               end;
               EOL := true;
               break;
             end;
             NextChar(charLen);
           end;
           if EOL then break;
         end;
  end;

  if s <> '' then
    AWordList.Add(s);

  if (part <> '') then
  begin
    if (FTextRotation <> rtStacked) then
      ALineWidth := ALineWidth + FCanvas.TextWidth(part)
    else
      ALineWidth := ALineWidth + tmpWidth;
  end;
end;

{ The scanner has reached the text character at the specified position.
  Determines the
  - index of the NEXT rich-text parameter (ANextRtParamIndex)
  - character index where NEXT font change will occur (ACharIndexOfNextFont)
  - CURRENT font height (ACurrFontHeight)
  - CURRENT font position (normal/sub/super) (ACurrFontPos) }
procedure TsTextPainter.UpdateFont(ACharIndex: Integer;
  var ACurrRtpIndex, ACharIndexOfNextFont, ACurrFontHeight: Integer;
  var ACurrFontPos: TsFontPosition);
var
  fnt: TsFont;
begin
  if (ACurrRtpIndex < High(FRtParams)) and (ACharIndex = ACharIndexOfNextFont) then
  begin
    inc(ACurrRtpIndex);
    if ACurrRtpIndex < High(FRtParams) then
      ACharIndexOfNextFont := FRtParams[ACurrRtpIndex+1].FirstIndex else
      ACharIndexOfNextFont := MaxInt;
    fnt := FWorkbook.GetFont(FRtParams[ACurrRtpIndex].FontIndex);
    Convert_sFont_to_Font(fnt, FCanvas.Font);
    ACurrFontHeight := FCanvas.TextHeight('Tg');
    if fnt.Position <> fpNormal then
      FCanvas.Font.Size := round(fnt.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR);
    ACurrFontPos := fnt.Position;
  end;
end;


end.
