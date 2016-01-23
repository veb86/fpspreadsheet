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
  ARotation: TsTextRotation; AOverrideTextColor: TColor; ARightToLeft: Boolean);

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean): Integer;

function RichTextHeight(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean): Integer;

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
    FStackPeriod: Integer;
    FLines: TFPList;
    // Scanner
    FRtpIndex: Integer;
    FCharIndex: integer;
    FCharIndexOfNextFont: Integer;
    FFontHeight: Integer;
    FFontPos: TsFontPosition;
    FPtr: PChar;
  private
    function GetHeight: Integer;
    function GetWidth: Integer;
  protected
    procedure DrawLine(pEnd: PChar; x, y, ALineHeight: Integer; AOverrideTextColor: TColor);
    procedure DrawText(var x, y: Integer; s: String; ALineHeight: Integer);
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
      AVertAlignment: TsVertAlignment; AWordWrap, ARightToLeft: Boolean);
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
                                                             (*
{------------------------------------------------------------------------------}
{                       Processing of rich-text                                }
{------------------------------------------------------------------------------}
type
  TLineInfo = record
    pStart, pEnd: PChar;
    Words: Array of String;
    NumSpaces: Integer;
    BeginsWithFontOfRtpIndex: Integer;
    Width: Integer;
    Height: Integer;
  end;


procedure InternalDrawRichText(ACanvas: TCanvas; AWorkbook: TsWorkbook;
  const ARect: TRect; const AText: String; AFontIndex: Integer;
  ARichTextParams: TsRichTextParams; AWordwrap: Boolean;
  AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation; AOverrideTextColor: TColor;
  ARightToLeft, AMeasureOnly: Boolean;
  var AWidth, AHeight: Integer);
var
  xpos, ypos: Integer;
  p: PChar;
  lRtpIndex: Integer;
  lLineInfo: TLineInfo;
  lLineInfos: array of TLineInfo = nil;
  lTotalHeight, lLinelen: Integer;
  lStackPeriod: Integer = 0;
  lCharPos: Integer;
  lFontPos: TsFontPosition;
  lFontHeight: Integer;
  lCharIndexFontChange: Integer;
  ts: TTextStyle;

  { Scans the line for a possible line break. The max width is determined by
    the size of the rectangle ARect passed to the outer procedure:
    rectangle width in case of horizontal painting, rectangle height in case
    of vertical painting. Line breaks can occure at spaces or cr/lf characters,
    or, if not found, at any character reaching the max width.

    Parameters:

    P              defines where the scan starts. At the end of the routine it
                   points to the first character of the next line.
    ANumSpaces     is how many spaces were found between the start and end value
                   of P.
    ARtpFontIndex  At input, this is the index of the rich-text formatting
                   parameter value used for the font at line start. At output,
                   it is the index which will be valid at next line start.
    ALineWidth     the pixel width of the line seen along drawing direction, i.e.
                   in case of stacked text it is the character height times
                   character count in the line (!)
    ALineHeight    The height of the line as seen vertically to the drawing
                   direction. Normally this is the height of the largest font
                   found in the line; in case of stacked text it is the
                   standardized width of a character. }
  procedure ScanLine(var P: PChar; var ALineInfo: TLineInfo;
    var ANextLineRtParamIndex: Integer);
  var
    pWordStart: PChar;
    EOL: Boolean;
    savedSpaces: Integer;
    savedWidth: Integer;
    savedCharPos: Integer;
//    savedRtpFontIndex: Integer;
    savedNextLineRtParamIndex: Integer;
    maxWidth: Integer;
    dw: Integer;
    lineChar: utf8String;
    charLen: Integer;    // Number of bytes of current utf8 character
    s: String;

              {
    TLineInfo = record
      pStart, pEnd: PChar;
      Words: Array of String;
      NumSpaces: Integer;
      BeginsWithFontOfRtpIndex: Integer;
      Width: Integer;
      Height: Integer;
    end;
               }
  begin
    ALineInfo.pStart := P;
    ALineInfo.pEnd := P;
    ALineInfo.NumSpaces := 0;
    ALineInfo.BeginsWithFontOfRtpIndex := ANextLineRtParamIndex;
    ALineInfo.Width := 0;
    ALineInfo.Height := lFontHeight;
    SetLength(ALineInfo.Words, 0);

    s := '';
    savedWidth := 0;
    savedSpaces := 0;
    maxWidth := MaxInt;
    if AWordwrap then
    begin
      if ARotation = trHorizontal then
        maxWidth := ARect.Right - ARect.Left
      else
        maxWidth := ARect.Bottom - ARect.Top;
    end;

    UpdateFont(ACanvas, AWorkbook, AFontIndex, ARichTextParams, lCharPos,
      ANextLineRtParamIndex, lFontHeight, lFontPos);
    ALineInfo.Height := Max(fontHeight, ALineInfo.Height);

    while P^ <> #0 do begin
      case P^ of
        #13: begin
               inc(P);
               inc(lCharPos);
               if P^ = #10 then
               begin
                 inc(P);
                 inc(lCharPos);
               end;
               break;
             end;
        #10: begin
               inc(P);
               inc(lCharPos);
               break;
             end;
        ' ': begin
               SetLength(ALineInfo.Words, Length(ALineInfo.Words)+1);
               ALineInfo.Words[High(ALineInfo.Words)] := s;
               savedWidth := ALineInfo.Width;
               savedSpaces := ALineInfo.NumSpaces;
               // Find next word
               while P^ = ' ' do
               begin
                 UpdateFont(ACanvas. AWorkbook, AFontIndex, ARichTextParams,
                   lCharPos, ANextLineRtParamIndex, lFontHeight, lFontPos);
                 ALineInfo.Height := Max(lFontHeight, ALineInfo.Height);
                 dw := Math.IfThen(ARotation = rtStacked, lFontHeight, ACanvas.TextWidth(' '));
                 AALineInfo.Width := ALineInfo.Width + dw;
                 inc(ALineInfo.NumSpaces);
                 inc(P);
                 inc(lCharPos);
               end;
               if ALineInfo.Width >= maxWidth then
               begin
                 ALineInfo.Width := savedWidth;
                 ALineInfo.NumSpaces := savedSpaces;
                 break;
               end;
             end;
        else begin
               // Bere begins a new word. Find end of this word and check if
               // it fits into the line.
               // Store the data valid for the word start.
               pWordStart := P;
               s := '';
               savedCharPos := lCharPos;
               savedNextLineTrParamIndex := ANextLineParamIndex;
               EOL := false;
               while (P^ <> #0) and (P^ <> #13) and (P^ <> #10) and (P^ <> ' ') do
               begin
                 UpdateFont(ACanvas, AWorkbook, AFontIndex, ARichTextParams,
                   lCharPos, ANextLineRtParamIndex, lFontHeight, lFontPos);
                 ALineInfo.Height := Max(lFontHeight, ALineInfo.Height);
                 lineChar := UnicodeToUTF8(UTF8CharacterToUnicode(p, charLen));
                 s := s + lineChar;
                 dw := Math.IfThen(ARotation = rtStacked, lFontHeight, ACanvas.TextWidth(lineChar));
                 ALineInfo.Width := ALineInfo.Width + dw;
                 if ALineInfo.Width > maxWidth then
                 begin
                   // The line exeeds the max line width.
                   // There are two cases:
                   if ALineInfo.NumSpaces > 0 then
                   begin
                     // (a) This is not the only word: Go back to where this
                     // word began. We had stored everything needed!
                     P := pWordStart;
                     lCharPos := savedCharPos;
                     ALineInfo.Width := savedWidth;
                     ANextLineParamIndex := savedNextLineParamIndex;
                   end;
                   // (b) This is the only word in the line --> we break at the
                   // current cursor position.
                   EOL := true;
                   break;
                 end;
                 inc(P);
                 inc(lCharPos);
               end;
               if EOL then break;
             end;
      end;
    end;
    UpdateFont(ACanvas, AWorkbook, AFontIndex
    UpdateFont(charPos, ARtpFontIndex, fontHeight, fontPos);
    ALineHeight := Max(fontHeight, ALineHeight);
  end;

  procedure DrawText(var x, y: Integer; ALineHeight: Integer; s: String);
  var
    w: Integer;
  begin
    w := ACanvas.TextWidth(s);

    case ARotation of
      trHorizontal:
        begin
          ACanvas.Font.Orientation := 0;
          if ARightToLeft then
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x-w, y, s);
              fpSubscript  : ACanvas.TextOut(x-w, y+ALineHeight div 2, s);
              fpSuperScript: ACanvas.TextOut(x-w, y-ALineHeight div 6, s);
            end;
            dec(x, w);
          end else
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y, s);
              fpSubscript  : ACanvas.TextOut(x, y+ALineHeight div 2, s);
              fpSuperscript: ACanvas.TextOut(x, y-ALineHeight div 6, s);
            end;
            inc(x, w);
          end;
        end;

      rt90DegreeClockwiseRotation:
        begin
          ACanvas.Font.Orientation := -900;
          if ARightToLeft then
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y-w, s);
              fpSubscript  : ACanvas.TextOut(x-ALineHeight div 2, y-w, s);
              fpSuperscript: ACanvas.TextOut(x+ALineHeight div 6, y-w, s);
            end;
            dec(y, w);
          end else
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y, s);
              fpSubscript  : ACanvas.TextOut(x-ALineHeight div 2, y, s);
              fpSuperscript: ACanvas.TextOut(x+ALineHeight div 6, y, s);
            end;
            inc(y, w);
          end;
        end;

      rt90DegreeCounterClockwiseRotation:
        begin
          ACanvas.Font.Orientation := +900;
          if ARightToLeft then
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y+w, s);
              fpSubscript  : ACanvas.TextOut(x+ALineHeight div 2, y+w, s);
              fpSuperscript: ACanvas.TextOut(x-ALineHeight div 6, y+w, s);
            end;
            inc(y, w);
          end else
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y, s);
              fpSubscript  : ACanvas.TextOut(x+ALineHeight div 2, y, s);
              fpSuperscript: ACanvas.TextOut(x-ALineHeight div 6, y, s);
            end;
            dec(y, w);
          end;
        end;

      rtStacked:
        begin
          ACanvas.Font.Orientation := 0;
          w := ACanvas.TextWidth(s);
          // chars centered around x
          if ARightToLeft then       // is this ok??
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x - w div 2, y-fontheight, s);
              fpSubscript  : ACanvas.TextOut(x - w div 2, y-fontheight+ALineHeight div 2, s);
              fpSuperscript: ACanvas.TextOut(x - w div 2, y-fontheight-ALineHeight div 6, s);
            end;
            dec(y, fontHeight);
          end else
          begin
            case fontpos of
              fpNormal     : ACanvas.TextOut(x - w div 2, y, s);
              fpSubscript  : ACanvas.TextOut(x - w div 2, y+ALineHeight div 2, s);
              fpSuperscript: ACanvas.TextOut(x - w div 2, y-ALineHeight div 6, s);
            end;
            inc(y, fontHeight);
          end;
        end;
    end;
  end;

  procedure DrawLine(pStart, pEnd: PChar; x, y, ALineHeight, ARtpFontIndex: Integer);
  var
    p: PChar;
    charPosForNextFont, charLen: Integer;
    s: String;
    fntIdx: Integer;
  begin
    p := pStart;
    s := '';
    charPosForNextFont := ARichTextParams[ARtpFontIndex].FirstIndex;
    while (p^ <> #0) and (p < pEnd) do begin
      case p^ of
        #10: begin
               DrawText(x, y, ALineHeight, s);
               s := '';
               inc(p);
               inc(charpos);
               break;
             end;
        #13: begin
               DrawText(x, y, ALineHeight, s);
               s := '';
               inc(p);
               inc(charpos);
               if p^ = #10 then
               begin
                 inc(p);
                 inc(charpos);
               end;
               break;
             end;
        else
             s := s + UnicodeToUTF8(UTF8CharacterToUnicode(p, charLen));
             if CharPos = charPosForNextFont then begin
               DrawText(x, y, ALineHeight, s);
               s := '';
             end;
             inc(charPos);
             inc(p, charLen);
             UpdateFont(charPos, ARtpFontIndex, fontheight, fontpos);
      end;
    end;
    if s <> '' then
      DrawText(x, y, ALineHeight, s);
  end;
                               (*
  { Paints the text between the pointers pStart and pEnd.
    Starting point for the text location is given by the coordinates x/y, i.e.
    text alignment is already corrected. In case of sub/superscripts, the
    characters reduced in size are shifted vertical to drawing direction by a
    fraction of the line height (ALineHeight).
    ARtpFontIndex is the index of the rich-text formatting param used to at line
    start for font selection; it will advance automatically along the line }
  procedure DrawLine(pStart, pEnd: PChar; x,y, ALineHeight: Integer;
    ARtpFontIndex: Integer);
  var
    p: PChar;
    w: Integer;
    s: utf8String;
    charLen: Integer;
  begin
    p := pStart;
    while p^ <> #0 do begin
      s := UnicodeToUTF8(UTF8CharacterToUnicode(p, charLen));
      UpdateFont(charPos, ARtpFontIndex, fontHeight, fontPos);
      if AOverrideTextColor <> clNone then
        ACanvas.Font.Color := AOverrideTextColor;
      case p^ of
        #10: begin
               inc(p);
               inc(charPos);
               break;
             end;
        #13: begin
               inc(p);
               inc(charPos);
               if p^ = #10 then begin
                 inc(p);
                 inc(charpos);
               end;
               break;
             end;
      end;
      case ARotation of
        trHorizontal:
          begin
            ACanvas.Font.Orientation := 0;
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y, s);
              fpSubscript  : ACanvas.TextOut(x, y + ALineHeight div 2, s);
              fpSuperscript: ACanvas.TextOut(x, y - ALineHeight div 6, s);
            end;
            inc(x, ACanvas.TextWidth(s));
          end;
        rt90DegreeClockwiseRotation:
          begin
            ACanvas.Font.Orientation := -900;
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y, s);
              fpSubscript  : ACanvas.TextOut(x - ALineHeight div 2, y, s);
              fpSuperscript: ACanvas.TextOut(x + ALineHeight div 6, y, s);
            end;
            inc(y, ACanvas.TextWidth(s));
          end;
        rt90DegreeCounterClockwiseRotation:
          begin
            ACanvas.Font.Orientation := +900;
            case fontpos of
              fpNormal     : ACanvas.TextOut(x, y, s);
              fpSubscript  : ACanvas.TextOut(x + ALineHeight div 2, y, s);
              fpSuperscript: ACanvas.TextOut(x - ALineHeight div 6, y, s);
            end;
            dec(y, ACanvas.TextWidth(s));
          end;
        rtStacked:
          begin
            ACanvas.Font.Orientation := 0;
            w := ACanvas.TextWidth(s);
            // chars centered around x
            case fontpos of
              fpNormal     : ACanvas.TextOut(x - w div 2, y, s);
              fpSubscript  : ACanvas.TextOut(x - w div 2, y + ALineHeight div 2, s);
              fpSuperscript: ACanvas.TextOut(x - w div 2, y - ALineHeight div 6, s);
            end;
            inc(y, fontHeight);
          end;
      end;

      inc(P, charLen);
      inc(charPos);
      if P >= PEnd then break;
    end;
    UpdateFont(charPos, ARtpFontIndex, fontHeight, fontPos);
  end;                           *)

begin
  if AText = '' then
    exit;

  p := PChar(AText);
  lCharPos := 1;      // Counter for utf8 character position
  lTotalHeight := 0;
  lLinelen := 0;

  ts := ACanvas.TextStyle;
  ts.RightToLeft := ARightToLeft;
  ACanvas.TextStyle := ts;

  // (1) Get layout of lines
  //  ======================
  // "lineinfos" collect data for where lines start and end, their width and
  // height, the rich-text parameter index range, and the number of spaces
  // (for text justification)
  InitFont(ACanvas, AWorkbook, AFontIndex, ARichTextParams, lRtpIndex,
    lCharIndexFontChange, lFontHeight, lFontPos);
  if ARotation = rtStacked then
    lStackPeriod := ACanvas.TextWidth('M') * 2;
  SetLength(lLineInfos, 0);
  repeat
    SetLength(lLineInfos, Length(lLineInfos)+1);
    with lLineInfos[High(lLineInfos)] do begin
      pStart := p;
      pEnd := p;
      BeginsWithFontOfRtpIndex := lRtpIndex;
      ScanLine(pStart, lLineInfos[High(lLineInfos)], pEnd, NumSpaces, rtpIndex, Width, Height);
      totalHeight := totalHeight + Height;
      linelen := Max(linelen, Width);
      p := pEnd;
    end;
  until p^ = #0;

  AWidth := linelen;
  if ARotation = rtStacked then
    AHeight := Length(lineinfos) * stackperiod  // to be understood horizontally
  else
    AHeight := totalHeight;
  if AMeasureOnly then
    exit;

  // (2) Draw lines
  // ==============
  // 2a) get starting point of line
  // ------------------------------
  case ARotation of
    trHorizontal:
      case AVertAlignment of
        vaTop   : ypos := ARect.Top;
        vaBottom: ypos := ARect.Bottom - totalHeight;
        vaCenter: ypos := (ARect.Top + ARect.Bottom - totalHeight) div 2;
      end;
    rt90DegreeClockwiseRotation:
      case AHorAlignment of
        haLeft  : xpos := ARect.Left + totalHeight;
        haRight : xpos := ARect.Right;
        haCenter: xpos := (ARect.Left + ARect.Right + totalHeight) div 2;
      end;
    rt90DegreeCounterClockwiseRotation:
      case AHorAlignment of
        haLeft  : xpos := ARect.Left;
        haRight : xpos := ARect.Right - totalHeight;
        haCenter: xpos := (ARect.Left + ARect.Right - totalHeight) div 2;
      end;
    rtStacked:
      begin
        totalHeight := (Length(lineinfos) - 1) * stackperiod;
        case AHorAlignment of
          haLeft  : xpos := ARect.Left + stackPeriod div 2;
          haRight : xpos := ARect.Right - totalHeight + stackPeriod div 2;
          haCenter: xpos := (ARect.Left + ARect.Right - totalHeight) div 2;
        end;
      end;
  end;

  // (2b) Draw line by line and respect text rotation
  // ------------------------------------------------
  charPos := 1;      // Counter for utf8 character position
  InitFont(rtpIndex, fontheight, fontpos);
  for lineInfo in lineInfos do begin
    with lineInfo do
    begin
      p := pStart;
      case ARotation of
        trHorizontal:
          begin
            if ARightToLeft then
              case AHorAlignment of
                haLeft   : xpos := ARect.Left + Width;
                haRight  : xpos := ARect.Right;
                haCenter : xpos := (ARect.Left + ARect.Right + Width) div 2;
              end
            else
              case AHorAlignment of
                haLeft   : xpos := ARect.Left;
                haRight  : xpos := ARect.Right - Width;
                haCenter : xpos := (ARect.Left + ARect.Right - Width) div 2;
              end;
            DrawLine(pStart, pEnd, xpos, ypos, Height, BeginsWithFontOfRtpIndex);
            inc(ypos, Height);
          end;
        rt90DegreeClockwiseRotation:
          begin
            if ARightToLeft then
              case AVertAlignment of
                vaTop    : ypos := ARect.Top + Width;
                vaBottom : ypos := ARect.Bottom;
                vaCenter : ypos := (ARect.Top + ARect.Bottom + Width) div 2;
              end
            else
              case AVertAlignment of
                vaTop    : ypos := ARect.Top;
                vaBottom : ypos := ARect.Bottom - Width;
                vaCenter : ypos := (ARect.Top + ARect.Bottom - Width) div 2;
              end;
            DrawLine(pStart, pEnd, xpos, ypos, Height, BeginsWithFontOfRtpIndex);
            dec(xpos, Height);
          end;
        rt90DegreeCounterClockwiseRotation:
          begin
            if ARightToLeft then
              case AVertAlignment of
                vaTop    : ypos := ARect.Top;
                vaBottom : ypos := ARect.Bottom - Width;
                vaCenter : ypos := (ARect.Top + ARect.Bottom - Width) div 2;
              end
            else
              case AVertAlignment of
                vaTop    : ypos := ARect.Top + Width;
                vaBottom : ypos := ARect.Bottom;
                vaCenter : ypos := (ARect.Top + ARect.Bottom + Width) div 2;
              end;
            DrawLine(pStart, pEnd, xpos, ypos, Height, BeginsWithFontOfRtpIndex);
            inc(xpos, Height);
          end;
        rtStacked:
          begin
            case AVertAlignment of
              vaTop    : ypos := ARect.Top;
              vaBottom : ypos := ARect.Bottom - Width;
              vaCenter : ypos := (ARect.Top + ARect.Bottom - Width) div 2;
            end;
            DrawLine(pStart, pEnd, xpos, ypos, Height, BeginsWithFontOfRtpIndex);
            inc(xpos, stackPeriod);
          end;
      end;
    end;
  end;
end;
                                                                 *)
procedure DrawRichText(ACanvas: TCanvas; AWorkbook: TsWorkbook; const ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  AWordwrap: Boolean; AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation; AOverrideTextColor: TColor; ARightToLeft: Boolean);
var
//  w: Integer = 0;
//  h: Integer = 0;
  painter: TsTextPainter;
begin
  painter := TsTextPainter.Create(ACanvas, AWorkbook, ARect, AText, ARichTextParams,
    AFontIndex, ARotation, AHorAlignment, AVertAlignment, AWordWrap, ARightToLeft);
  try
    painter.Draw(AOverrideTextColor);
  finally
    painter.Free;
  end;
      {
  InternalDrawRichText(ACanvas, AWorkbook, ARect, AText, AFontIndex,
    ARichTextParams, AWordWrap, AHorAlignment, AVertAlignment, ARotation,
    AOverrideTextColor, ARightToLeft, false, w, h);
    }
end;

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean): Integer;
var
//  h: Integer = 0;
//  w: Integer = 0;
  painter: TsTextPainter;
begin
  painter := TsTextPainter.Create(ACanvas, AWorkbook, ARect, AText, ARichTextParams,
    AFontIndex, ATextRotation, haLeft, vaTop, AWordWrap, ARightToLeft);
  try
    Result := painter.Height;
  finally
    painter.Free;
  end;
                            {
  InternalDrawRichText(ACanvas, AWorkbook, AMaxRect, AText, AFontIndex,
    ARichTextParams, AWordWrap, haLeft, vaTop, ATextRotation, clNone,
    ARightToLeft, true, w, h
  );
  case ATextRotation of
    trHorizontal, rtStacked:
      Result := w;
    rt90DegreeClockwiseRotation, rt90DegreeCounterClockwiseRotation:
      Result := h;
  end;                       }
end;

function RichTextHeight(ACanvas: TCanvas; AWorkbook: TsWorkbook; ARect: TRect;
  const AText: String; ARichTextParams: TsRichTextParams; AFontIndex: Integer;
  ATextRotation: TsTextRotation; AWordWrap, ARightToLeft: Boolean): Integer;
var
  painter: TsTextPainter;
//  h: Integer = 0;
//  w: Integer = 0;
begin
  painter := TsTextPainter.Create(ACanvas, AWorkbook, ARect, AText, ARichTextParams,
    AFontIndex, ATextRotation, haLeft, vaTop, AWordWrap, ARightToLeft);
  try
    Result := painter.Height;
  finally
    painter.Free;
  end;

  {
  InternalDrawRichText(ACanvas, AWorkbook, AMaxRect, AText, AFontIndex,
    ARichTextParams, AWordWrap, haLeft, vaTop, ATextRotation, clNone,
    ARightToLeft, true, w, h
  );
  case ATextRotation of
    trHorizontal:
      Result := h;
    rt90DegreeClockwiseRotation, rt90DegreeCounterClockwiseRotation, rtStacked:
      Result := w;
  end;
  }
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
  AVertAlignment: TsVertAlignment; AWordWrap, ARightToLeft: Boolean);
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
  Prepare;
end;

destructor TsTextPainter.Destroy;
var
  j: Integer;
begin
  for j:=FLines.Count-1 downto 0 do TObject(FLines[j]).Free;
  FLines.Free;
  inherited Destroy;
end;

{ Draw the lines }
procedure TsTextPainter.Draw(AOverrideTextColor: TColor);
var
  xpos, ypos: Integer;
  totalHeight: Integer;
  lineinfo: TsLineInfo;
  pEnd: PChar;
  j: Integer;
begin
  // (1) Get starting point of line
  case FTextRotation of
    trHorizontal:
      case FVertAlignment of
        vaTop   : ypos := FRect.Top;
        vaBottom: ypos := FRect.Bottom - FTotalHeight;
        vaCenter: ypos := (FRect.Top + FRect.Bottom - FTotalHeight) div 2;
      end;
    rt90DegreeClockwiseRotation:
      case FHorAlignment of
        haLeft  : xpos := FRect.Left + FTotalHeight;
        haRight : xpos := FRect.Right;
        haCenter: xpos := (FRect.Left + FRect.Right + FTotalHeight) div 2;
      end;
    rt90DegreeCounterClockwiseRotation:
      case FHorAlignment of
        haLeft  : xpos := FRect.Left;
        haRight : xpos := FRect.Right - FTotalHeight;
        haCenter: xpos := (FRect.Left + FRect.Right - FTotalHeight) div 2;
      end;
    rtStacked:
      begin
        totalHeight := (FLines.Count - 1) * FStackperiod;
        case FHorAlignment of
          haLeft  : xpos := FRect.Left + FStackPeriod div 2;
          haRight : xpos := FRect.Right - totalHeight + FStackPeriod div 2;
          haCenter: xpos := (FRect.Left + FRect.Right - totalHeight) div 2;
        end;
      end;
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
    with lineInfo do
    begin
      case FTextRotation of
        trHorizontal:
          begin
            if FRightToLeft then
              case FHorAlignment of
                haLeft   : xpos := FRect.Left + Width;
                haRight  : xpos := FRect.Right;
                haCenter : xpos := (FRect.Left + FRect.Right + Width) div 2;
              end
            else
              case FHorAlignment of
                haLeft   : xpos := FRect.Left;
                haRight  : xpos := FRect.Right - Width;
                haCenter : xpos := (FRect.Left + FRect.Right - Width) div 2;
              end;
            DrawLine(pEnd, xpos, ypos, Height, AOverrideTextColor);
            inc(ypos, Height);
          end;
        rt90DegreeClockwiseRotation:
          begin
            if FRightToLeft then
              case FVertAlignment of
                vaTop    : ypos := FRect.Top + Width;
                vaBottom : ypos := FRect.Bottom;
                vaCenter : ypos := (FRect.Top + FRect.Bottom + Width) div 2;
              end
            else
              case FVertAlignment of
                vaTop    : ypos := FRect.Top;
                vaBottom : ypos := FRect.Bottom - Width;
                vaCenter : ypos := (FRect.Top + FRect.Bottom - Width) div 2;
              end;
            DrawLine(pEnd, xpos, ypos, Height, AOverrideTextColor);
            dec(xpos, Height);
          end;
        rt90DegreeCounterClockwiseRotation:
          begin
            if FRightToLeft then
              case FVertAlignment of
                vaTop    : ypos := FRect.Top;
                vaBottom : ypos := FRect.Bottom - Width;
                vaCenter : ypos := (FRect.Top + FRect.Bottom - Width) div 2;
              end
            else
              case FVertAlignment of
                vaTop    : ypos := FRect.Top + Width;
                vaBottom : ypos := FRect.Bottom;
                vaCenter : ypos := (FRect.Top + FRect.Bottom + Width) div 2;
              end;
            DrawLine(pEnd, xpos, ypos, Height, AOverrideTextColor);
            inc(xpos, Height);
          end;
        rtStacked:
          begin
            case FVertAlignment of
              vaTop    : ypos := FRect.Top;
              vaBottom : ypos := FRect.Bottom - Width;
              vaCenter : ypos := (FRect.Top + FRect.Bottom - Width) div 2;
            end;
            DrawLine(pEnd, xpos, ypos, Height, AOverrideTextColor);
            inc(xpos, FStackPeriod);
          end;
      end;
    end;
  end;

end;

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
  w: Integer;
  P: TPoint;
begin
  w := FCanvas.TextWidth(s);
  P := GetTextPt(x, y, ALineHeight);
  FCanvas.Font.Orientation := TEXT_ANGLE[FTextRotation];
  case FTextRotation of
    trHorizontal:
      begin
        if FRightToLeft
          then FCanvas.TextOut(P.x-w, P.y, s)
          else FCanvas.TextOut(P.x, P.y, s);
        inc(x, w*MULTIPLIER[FTextRotation, FRightToLeft]);
      end;
    rt90DegreeClockwiseRotation:
      begin
        if FRightToLeft
          then FCanvas.TextOut(P.x, P.y-w, s)
          else FCanvas.TextOut(P.x, p.y, s);
        inc(y, w*MULTIPLIER[FTextRotation, FRightToLeft]);
      end;
    rt90DegreeCounterClockwiseRotation:
      begin
        if FRightToLeft
          then FCanvas.TextOut(P.x, P.y+w, s)
          else FCanvas.TextOut(P.x, P.y, s);
        inc(y, w*MULTIPLIER[FTextRotation, FRightToLeft]);
      end;
    rtStacked:
      begin                       // IS THIS OK?
        w := FCanvas.TextWidth(s);
        // chars centered around x
        if FRightToLeft
          then FCanvas.TextOut(P.x - w div 2, P.y - FFontHeight, s)
          else FCanvas.TextOut(P.x - w div 2, P.y, s);
        inc(y, FFontHeight * MULTIPLIER[FTextRotation, FRightToLeft]);
      end;
  end;
end;

function TsTextPainter.GetHeight: Integer;
begin
  if FTextRotation = rtStacked then
    Result := FLines.Count * FStackperiod  // to be understood horizontally
  else
    Result := FTotalHeight;
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
  Result := FMaxLineLen;
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
  ACurrFontHeight := FCanvas.TextHeight('Tg');
  if (fnt <> nil) and (fnt.Position <> fpNormal) then
    FCanvas.Font.Size := round(fnt.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR);
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
  if FTextRotation = rtStacked then
    FStackPeriod := FCanvas.TextWidth('M') * 2;

  FPtr := PChar(FText);
  FCharIndex := 1;
  while (FPtr^ <> #0) do begin
    lineInfo := TsLineInfo.Create;
    lineInfo.pStart := FPtr;
    lineInfo.BeginsWithFontOfRtpIndex := FRtpIndex;
    ScanLine(lineInfo.NumSpaces, lineInfo.Width, lineInfo.Height, lineInfo.WordList);
    FLines.Add(lineinfo);
    FTotalHeight := FTotalHeight + lineInfo.Height;
    FMaxLineLen := Max(FMaxLineLen, lineInfo.Width);
  end;
end;

{ Scans the line for a possible line break. The max width is determined by
  the size of the rectangle ARect passed to the outer procedure:
  rectangle width in case of horizontal painting, rectangle height in case
  of vertical painting. Line breaks can occure at spaces or cr/lf characters,
  or, if not found, at any character reaching the max width.

  Parameters:

  P              defines where the scan starts. At the end of the routine it
                 points to the first character of the next line.
  ANumSpaces     is how many spaces were found between the start and end value
                 of P.
  ARtpFontIndex  At input, this is the index of the rich-text formatting
                 parameter value used for the font at line start. At output,
                 it is the index which will be valid at next line start.
  ALineWidth     the pixel width of the line seen along drawing direction, i.e.
                 in case of stacked text it is the character height times
                 character count in the line (!)
  ALineHeight    The height of the line as seen vertical to the drawing
                 direction. Normally this is the height of the largest font
                 found in the line; in case of stacked text it is the
                 standardized width of a character. }
procedure TsTextPainter.ScanLine(var ANumSpaces, ALineWidth, ALineHeight: Integer;
  AWordList: TStringList);
var
  tmpWidth: Integer;
  savedWidth: Integer;
  savedSpaces: Integer;
  savedCharIndex: Integer;
  savedCurrRtpIndex: Integer;
  savedCharIndexOfNextFont: Integer;
  maxWidth: Integer;
  s: String;
  charLen: Integer;
  ch: String;
  dw: Integer;
  EOL: Boolean;
  pWordStart: PChar;
  part: String;
  savedpart: String;
  PStart: PChar;
begin
  ANumSpaces := 0;
  ALineHeight := FFontHeight;
  ALineWidth := 0;
  savedWidth := 0;
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

  PStart := FPtr;
  while (FPtr^ <> #0) do
  begin
    case FPtr^ of
      #13: begin
             if (part <> '') and (FTextRotation <> rtStacked) then
               ALineWidth := ALineWidth + FCanvas.TextWidth(part);
             part := '';
             NextChar(1);
             if FPtr^ = #10 then
               NextChar(1);
             break;
           end;
      #10: begin
             if (part <> '') and (FTextRotation <> rtStacked) then
               ALineWidth := ALineWidth + FCanvas.TextWidth(part);
             part := '';
             NextChar(1);
             break;
           end;
      ' ': begin
             AWordList.Add(s);
             savedWidth := ALineWidth;
             savedSpaces := ANumSpaces;
             // Find next word
             while FPtr^ = ' ' do
             begin
               if (FCharIndex = FCharIndexOfNextFont) then
               begin
                 if (FTextRotation <> rtStacked) then
                   ALineWidth := ALineWidth + FCanvas.TextWidth(part);
                 part := '';
               end;
               UpdateFont(FCharIndex, FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
               if FTextRotation = rtStacked then
                 ALineWidth := ALineWidth + FFontHeight else
                 part := part + ' ';
               ALineHeight := Max(FFontHeight, ALineHeight);
               inc(ANumSpaces);
               NextChar(1);
             end;
             if ALineWidth >= maxWidth then
             begin
               ALineWidth := savedWidth;
               ANumSpaces := savedSpaces;
               part := '';
               break;
              end;
           end;
      else
           // Here, a new word begins. Find the end of this word and check if
           // it fits into the line.
           // Store the data valid for the word start.
           s := '';
           pWordStart := FPtr;
           savedCharIndex := FCharIndex;
           savedCurrRtpIndex := FRtpIndex;
           savedCharIndexOfNextFont := FCharIndexOfNextFont;
           savedpart := part;
           tmpWidth := 0;
           EOL := false;
           while (FPtr^ <> #0) and (FPtr^ <> #13) and (FPtr^ <> #10) and (FPtr^ <> ' ') do
           begin
             if FCharIndex = FCharIndexOfNextFont then
             begin
               if (FTextRotation <> rtStacked) then
                 ALineWidth := ALineWidth + FCanvas.TextWidth(part);
               part := '';
             end;
             UpdateFont(FCharIndex, FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
             ch := UnicodeToUTF8(UTF8CharacterToUnicode(FPtr, charLen));
             part := part + ch;
             tmpWidth := IfThen(FTextRotation = rtStacked, tmpWidth + FFontHeight, FCanvas.TextWidth(part));
             if ALineWidth + tmpWidth <= maxWidth then
             begin
               s := s + ch;
               ALineHeight := Max(FFontHeight, ALineHeight);
             end else
             begin
               // The line exeeds the max line width.
               // There are two cases:
               if ANumSpaces > 0 then
               begin
                 // (a) This is not the only word: Go back to where this
                 // word began. We had stored everything needed!
                 FPtr := pWordStart;
                 FCharIndex := savedCharIndex;
                 FCharIndexOfNextFont := savedCharIndexOfNextFont;
                 FRtpIndex := savedCurrRtpIndex;
                 part := '';
               end else
               begin
                 // (b) This is the only word in the line --> we break at the
                 // current cursor position.
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

  if (part <> '') and (FTextRotation <> rtStacked) then
    ALineWidth := ALineWidth + FCanvas.TextWidth(part);

  UpdateFont(FCharIndex, FRtpIndex, FCharIndexOfNextFont, FFontHeight, FFontPos);
  ALineHeight := Max(FFontHeight, ALineHeight);
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
