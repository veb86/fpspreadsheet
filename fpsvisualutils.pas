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
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  AWordwrap: Boolean; AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation; AOverrideTextColor: TColor);

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; AMaxRect: TRect;
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  ATextRotation: TsTextRotation; AWordWrap: Boolean): Integer;

function RichTextHeight(ACanvas: TCanvas; AWorkbook: TsWorkbook; AMaxRect: TRect;
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  ATextRotation: TsTextRotation; AWordWrap: Boolean): Integer;

{
function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; const AText: String;
  AFontIndex: Integer; ARichTextParams: TsRichTextParams): Integer;
}

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

procedure InternalDrawRichText(ACanvas: TCanvas; AWorkbook: TsWorkbook;
  const ARect: TRect; const AText: String; AFontIndex: Integer;
  ARichTextParams: TsRichTextParams; AWordwrap: Boolean;
  AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation; AOverrideTextColor: TColor; AMeasureOnly: Boolean;
  var AWidth, AHeight: Integer);
type
  TLineInfo = record
    pStart, pEnd: PChar;
    NumSpaces: Integer;
    BeginsWithFontOfRtpIndex: Integer;
    Width: Integer;
    Height: Integer;
  end;
var
  xpos, ypos: Integer;
  p, pStartText: PChar;
  rtpIndex: Integer;
  lineInfo: TLineInfo;
  lineInfos: Array of TLineInfo = nil;
  totalHeight, linelen, stackPeriod: Integer;
  charPos: Integer;
  fontpos: TsFontPosition;
  fontHeight: Integer;

  procedure InitFont(out ARtpFontIndex: Integer; out AFontHeight: Integer;
    out AFontPos: TsFontPosition);
  var
    fnt: TsFont;
    rtParam: TsRichTextParam;
  begin
    if (Length(ARichTextParams) > 0) and (charPos >= ARichTextParams[0].FirstIndex) then
    begin
      ARtpFontIndex := 0;
      fnt := AWorkbook.GetFont(ARichTextParams[0].FontIndex);
    end else
    begin
      ARtpFontIndex := -1;
      fnt := AWorkbook.GetFont(AFontIndex);
    end;
    Convert_sFont_to_Font(fnt, ACanvas.Font);
    AFontHeight := ACanvas.TextHeight('Tg');
    if (fnt <> nil) and (fnt.Position <> fpNormal) then
      ACanvas.Font.Size := round(fnt.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR);
    AFontPos := fnt.Position;
  end;

  procedure UpdateFont(ACharPos: Integer; var ARtpFontIndex: Integer;
    var AFontHeight: Integer; var AFontPos: TsFontPosition);
  var
    rtParam: TsRichTextParam;
    fnt: TsFont;
    endPos: Integer;
  begin
    if ARtpFontIndex = High(ARichTextParams) then
      endPos := MaxInt
    else begin
      rtParam := ARichTextParams[ARtpFontIndex + 1];
      endPos := rtParam.FirstIndex;
    end;

    if ACharPos >= endPos then begin
      inc(ARtpFontIndex);
      rtParam := ARichTextParams[ARtpFontIndex];
      fnt := AWorkbook.GetFont(rtParam.FontIndex);
      Convert_sFont_to_Font(fnt, ACanvas.Font);
      AFontHeight := ACanvas.TextHeight('Tg');
      if fnt.Position <> fpNormal then
        ACanvas.Font.Size := round(fnt.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR);
      AFontPos := fnt.Position;
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
  procedure ScanLine(var P: PChar; var NumSpaces: Integer;
    var ARtpFontIndex: Integer; var ALineWidth, ALineHeight: Integer);
  var
    pEOL: PChar;
    savedSpaces: Integer;
    savedWidth: Integer;
    savedRtpIndex: Integer;
    maxWidth: Integer;
    dw: Integer;
    spaceFound: Boolean;
    s: utf8String;
    charLen: Integer;    // Number of bytes of current utf8 character
  begin
    NumSpaces := 0;

    ALineHeight := fontHeight;
    ALineWidth := 0;
    savedWidth := 0;
    savedSpaces := 0;
    savedRtpIndex := ARtpFontIndex;
    spaceFound := false;
    pEOL := p;

    if AWordwrap then
    begin
      if ARotation = trHorizontal then
        maxWidth := ARect.Right - ARect.Left
      else
        maxWidth := ARect.Bottom - ARect.Top;
    end
    else
      maxWidth := MaxInt;

    while p^ <> #0 do begin
      UpdateFont(charPos, ARtpFontIndex, fontHeight, fontPos);
      ALineHeight := Max(fontHeight, ALineHeight);

      s := UnicodeToUTF8(UTF8CharacterToUnicode(p, charLen));
      case p^ of
        ' ': begin
               spaceFound := true;
               pEOL := p;
               savedWidth := ALineWidth;
               savedSpaces := NumSpaces;
               savedRtpIndex := ARtpFontIndex;
               dw := Math.IfThen(ARotation = rtStacked, fontHeight, ACanvas.TextWidth(s));
               if ALineWidth + dw < MaxWidth then
               begin
                 inc(NumSpaces);
                 ALineWidth := ALineWidth + dw;
               end else
                 break;
             end;
        #13: begin
               inc(p);
               inc(charPos);
               if p^ = #10 then
               begin
                 inc(p);
                 inc(charPos);
               end;
               break;
             end;
        #10: begin
               inc(p);
               inc(charPos);
               break;
             end;
        else begin
               dw := Math.IfThen(ARotation = rtStacked, fontHeight, ACanvas.TextWidth(s));
               ALineWidth := ALineWidth + dw;
               if ALineWidth > maxWidth then
               begin
                 if spaceFound then
                 begin
                   p := pEOL;
                   ALineWidth := savedWidth;
                   NumSpaces := savedSpaces;
                   ARtpFontIndex := savedRtpIndex;
                 end else
                 begin
                   ALineWidth := ALineWidth - dw;
                   if ALineWidth = 0 then
                     inc(p);
                 end;
                 break;
               end;
             end;
      end;

      inc(p, charLen);
      inc(charPos);
    end;
    UpdateFont(charPos, ARtpFontIndex, fontHeight, fontPos);
  end;

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
  end;

begin
  if AText = '' then
    exit;

  p := PChar(AText);
  pStartText := p;   // first char of text
  charPos := 1;      // Counter for utf8 character position
  totalHeight := 0;
  linelen := 0;

  Convert_sFont_to_Font(AWorkbook.GetFont(AFontIndex), ACanvas.Font);
  if ARotation = rtStacked then
    stackPeriod := ACanvas.TextWidth('M') * 2;

  // (1) Get layout of lines
  //  ======================
  // "lineinfos" collect data for where lines start and end, their width and
  // height, the rich-text parameter index range, and the number of spaces
  // (for text justification)
  InitFont(rtpIndex, fontheight, fontpos);
  repeat
    SetLength(lineInfos, Length(lineInfos)+1);
    with lineInfos[High(lineInfos)] do begin
      pStart := p;
      pEnd := p;
      BeginsWithFontOfRtpIndex := rtpIndex;
      ScanLine(pEnd, NumSpaces, rtpIndex, Width, Height);
      totalHeight := totalHeight + Height;
      linelen := Max(linelen, Width);
      p := pEnd;
      if p^ = ' ' then
        while (p^ <> #0) and (p^ = ' ') do begin
          inc(p);
          inc(charPos);
        end;
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

procedure DrawRichText(ACanvas: TCanvas; AWorkbook: TsWorkbook; const ARect: TRect;
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  AWordwrap: Boolean; AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation; AOverrideTextColor: TColor);
var
  w,h: Integer;
begin
  InternalDrawRichText(ACanvas, AWorkbook, ARect, AText, AFontIndex,
    ARichTextParams, AWordWrap, AHorAlignment, AVertAlignment, ARotation,
    AOverrideTextColor, false, w, h);
end;

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; AMaxRect: TRect;
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  ATextRotation: TsTextRotation; AWordWrap: Boolean): Integer;
var
  h, w: Integer;
begin
  InternalDrawRichText(ACanvas, AWorkbook, AMaxRect, AText, AFontIndex,
    ARichTextParams, AWordWrap, haLeft, vaTop, ATextRotation, clNone, true,
    w, h);
  case ATextRotation of
    trHorizontal, rtStacked:
      Result := w;
    rt90DegreeClockwiseRotation, rt90DegreeCounterClockwiseRotation:
      Result := h;
  end;
end;

function RichTextHeight(ACanvas: TCanvas; AWorkbook: TsWorkbook; AMaxRect: TRect;
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  ATextRotation: TsTextRotation; AWordWrap: Boolean): Integer;
var
  h, w: Integer;
begin
  InternalDrawRichText(ACanvas, AWorkbook, AMaxRect, AText, AFontIndex,
    ARichTextParams, AWordWrap, haLeft, vaTop, ATextRotation, clNone, true,
    w, h);
  case ATextRotation of
    trHorizontal:
      Result := h;
    rt90DegreeClockwiseRotation, rt90DegreeCounterClockwiseRotation, rtStacked:
      Result := w;
  end;
end;
            (*
function GetRichTextExtent(ACanvas: TCanvas; AWorkbook: TsWorkbook;
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  ATextRotation: TsTextRotation): TSize;
var
  s: String;
  p: Integer;
  len, height: Integer;
  rtp, next_rtp: TsRichTextParam;
  fnt, fnt0: TsFont;
begin
  Result := 0;
  if (ACanvas=nil) or (AWorkbook=nil) or (AText = '') then exit;

  fnt0 := AWorkbook.GetFont(AFontIndex);
  Convert_sFont_to_Font(fnt0, ACanvas.Font);

  if Length(ARichTextParams) = 0 then
  begin
    Result := ACanvas.TextExtent(AText);
    if ATextRotation = trHorizontal then
      exit;
    len := Result.cx;
    height := Result.cy;
    case ATextRotation of
      rt90DegreeClockwiseRotation,
      rt90DegreeCounterClockwiseRotation:
        begin
          Result.CX := height;
          Result.CY := len;
        end;
      rtStacked:
        begin
          Result.CX := ACanvas.TextWidth('M');
          Restul.CY := UTF8Length(AText) * height;
        end;
    end;
    exit;
  end;

  // Part with normal font before first rich-text parameter element
  rtp := ARichTextParams[0];
  if rtp.StartIndex > 0 then begin
    s := copy(AText, 1, rtp.StartIndex+1);  // StartIndex is 0-based
    Result := ACanvas.TextWidth(s);
    if fnt0.Position <> fpNormal then
      Result := Round(Result * SUBSCRIPT_SUPERSCRIPT_FACTOR);
  end;

  p := 0;
  while p < Length(ARichTextParams) do
  begin
    // Part with rich-text font
    rtp := ARichTextParams[p];
    fnt := AWorkbook.GetFont(rtp.FontIndex);
    Convert_sFont_to_Font(fnt, ACanvas.Font);
    s := copy(AText, rtp.StartIndex+1, rtp.EndIndex-rtp.StartIndex);
    w := ACanvas.TextWidth(s);
    if fnt.Position <> fpNormal then
      w := Round(w * SUBSCRIPT_SUPERSCRIPT_FACTOR);
    Result := Result + w;
    // Part with normal font
    if (p < High(ARichTextParams)-1) then
    begin
      next_rtp := ARichTextParams[p+1];
      n := next_rtp.StartIndex - rtp.EndIndex;
      if n > 0 then
      begin
        Convert_sFont_to_Font(fnt0, ACanvas.Font);
        s := Copy(AText, rtp.EndIndex, n);
        w := ACanvas.TextWidth(s);
        if fnt0.Position <> fpNormal then
          w := Round(w * SUBSCRIPT_SUPERSCRIPT_FACTOR);
        Result := Result + w;
      end;
    end;
    inc(p);
  end;
end;
              *)
end.
