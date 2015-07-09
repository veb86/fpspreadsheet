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
  ARotation: TsTextRotation);

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; const AText: String;
  AFontIndex: Integer; ARichTextParams: TsRichTextParams): Integer;


implementation

uses
  Types, Math, LCLType, LCLIntf, LazUTF8, fpsUtils;

const
{@@ Font size factor for sub-/superscript characters }
  SUBSCRIPT_SUPERSCRIPT_FACTOR = 0.6;

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

procedure DrawRichText(ACanvas: TCanvas; AWorkbook: TsWorkbook; const ARect: TRect;
  const AText: String; AFontIndex: Integer; ARichTextParams: TsRichTextParams;
  AWordwrap: Boolean; AHorAlignment: TsHorAlignment; AVertAlignment: TsVertAlignment;
  ARotation: TsTextRotation);
type
  TLineInfo = record
    pStart, pEnd: PChar;
    NumSpaces: Integer;
    FirstRtpIndex: Integer;
    NextRtpIndex: Integer;
    Width: Integer;
    Height: Integer;
  end;
  TRtState = (rtEnter, rtExit);
var
  xpos, ypos: Integer;
  p, pStartText: PChar;
  iRtp: Integer;
  lineInfo: TLineInfo;
  lineInfos: Array of TLineInfo = nil;
  totalHeight, stackPeriod: Integer;

  procedure InitFont(P: PChar; out rtState: TRtState;
    PendingRtpIndex: Integer; out AHeight: Integer);
  var
    fnt: TsFont;
    hasRtp: Boolean;
    rtp: TsRichTextParam;
  begin
    fnt := AWorkbook.GetFont(AFontIndex);
    hasRtp := PendingRtpIndex >= 0;
    if hasRTP and (PendingRtpIndex < Length(ARichTextParams)) then begin
      rtp := ARichTextParams[PendingRtpIndex];
      if p - pStartText >= rtp.StartIndex then
      begin
        fnt := AWorkbook.GetFont(rtp.FontIndex);
        rtState := rtEnter;
      end else
        rtState := rtExit;
    end;
    Convert_sFont_to_Font(fnt, ACanvas.Font);
    AHeight := ACanvas.TextHeight('Tg');
    if (fnt <> nil) and (fnt.Position <> fpNormal) then
      ACanvas.Font.Size := round(ACanvas.Font.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR);
  end;

  procedure UpdateFont(P:PChar; var rtState: TRtState;
    var PendingRtpIndex: Integer; var AHeight: Integer;
    out AFontPos: TsFontPosition);
  var
    hasRtp: Boolean;
    rtp: TsRichTextParam;
    fnt: TsFont;
  begin
    fnt := AWorkbook.GetFont(AFontIndex);
    hasRtp := PendingRtpIndex >= 0;
    if hasRtp and (PendingRtpIndex < Length(ARichTextParams)) then
    begin
      rtp := ARichTextParams[PendingRtpIndex];
      if (p - pStartText >= rtp.StartIndex) and (rtState = rtExit) then
      begin
        fnt := AWorkbook.GetFont(rtp.FontIndex);
        Convert_sFont_to_Font(fnt, ACanvas.Font);
        AHeight := ACanvas.TextHeight('Tg');
        if fnt.Position <> fpNormal then
          ACanvas.Font.Size := round(ACanvas.Font.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR);
        rtState := rtEnter;
      end else
      if (p - pStartText >= rtp.EndIndex) and (rtState = rtEnter) then
      begin
        inc(PendingRtpIndex);
        if PendingRtpIndex = Length(ARichTextparams) then
        begin
          fnt := AWorkbook.GetFont(AFontIndex);
          rtState := rtExit;
        end else
        begin
          rtp := ARichTextParams[PendingRtpIndex];
          if (p - pStartText < rtp.StartIndex) then
          begin
            fnt := AWorkbook.GetFont(AFontIndex);
            rtState := rtExit;
          end else
          begin
            fnt := AWorkbook.GetFont(rtp.FontIndex);
            rtState := rtEnter;
          end;
        end;
        Convert_sFont_to_Font(fnt, ACanvas.Font);
        AHeight := ACanvas.TextHeight('Tg');
        if fnt.Position <> fpNormal then
          ACanvas.Font.Size := round(ACanvas.Font.Size * SUBSCRIPT_SUPERSCRIPT_FACTOR);
      end;
    end;
    AFontPos := fnt.Position;
  end;

  procedure ScanLine(var P: PChar; var NumSpaces: Integer;
    var PendingRtpIndex: Integer; var width, height: Integer);
  var
    ch: Char;
    pEOL: PChar;
    savedSpaces: Integer;
    savedWidth: Integer;
    savedRtpIndex: Integer;
    maxWidth: Integer;
    rtState: TRtState;
    dw, h: Integer;
    fntpos: TsFontPosition;
    spaceFound: Boolean;
  begin
    NumSpaces := 0;

    InitFont(p, rtState, PendingRtpIndex, h);
    height := h;

    pEOL := p;
    width := 0;
    savedWidth := 0;
    savedSpaces := 0;
    savedRtpIndex := PendingRtpIndex;
    spaceFound := false;
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
      UpdateFont(p, rtState, PendingRtpIndex, h, fntpos);
      if h > height then height := h;

      ch := p^;
      case ch of
        ' ': begin
               spaceFound := true;
               pEOL := p;
               savedWidth := width;
               savedSpaces := NumSpaces;
               savedRtpIndex := PendingRtpIndex;
               dw := Math.IfThen(ARotation = rtStacked, h, ACanvas.TextWidth(ch));
               if width + dw < MaxWidth then
               begin
                 inc(NumSpaces);
                 width := width + dw;
               end else
                 break;
             end;
        #13,
        #10: begin
               dec(p);
               width := savedWidth;
               numSpaces := savedspaces;
               PendingRtpIndex := savedRtpIndex;
               exit;
             end;
        else begin
               dw := Math.IfThen(ARotation = rtStacked, h, ACanvas.TextWidth(ch));
               width := width + dw;
               if width > maxWidth then
               begin
                 if spaceFound then
                 begin
                   p := pEOL;
                   width := savedWidth;
                   NumSpaces := savedSpaces;
                   PendingRtpIndex := savedRtpIndex;
                 end else
                 begin
                   width := width - dw;
                   if width = 0 then
                     inc(p);
                 end;
                 break;
               end;
             end;
      end;

      inc(P, UTF8CharacterLength(p));
    end;
  end;

  procedure DrawLine(pStart, pEnd: PChar; x,y, hLine: Integer; PendingRtpIndex: Integer);
  var
    ch: Char;
    p: PChar;
    rtState: TRtState;
    h, w: Integer;
    fntpos: TsFontPosition;
  begin
    p := pStart;
    InitFont(p, rtState, PendingRtpIndex, h);
    while p^ <> #0 do begin
      UpdateFont(p, rtState, PendingRtpIndex, h, fntpos);
      ch := p^;
      case ARotation of
        trHorizontal:
          begin
            ACanvas.Font.Orientation := 0;
            case fntpos of
              fpNormal     : ACanvas.TextOut(x, y, ch);
              fpSubscript  : ACanvas.TextOut(x, y + hLine div 2, ch);
              fpSuperscript: ACanvas.TextOut(x, y - hLine div 6, ch);
            end;
            inc(x, ACanvas.TextWidth(ch));
          end;
        rt90DegreeClockwiseRotation:
          begin
            ACanvas.Font.Orientation := -900;
            case fntpos of
              fpNormal     : ACanvas.TextOut(x, y, ch);
              fpSubscript  : ACanvas.TextOut(x - hLine div 2, y, ch);
              fpSuperscript: ACanvas.TextOut(x + hLine div 6, y, ch);
            end;
            inc(y, ACanvas.TextWidth(ch));
          end;
        rt90DegreeCounterClockwiseRotation:
          begin
            ACanvas.Font.Orientation := +900;
            case fntpos of
              fpNormal     : ACanvas.TextOut(x, y, ch);
              fpSubscript  : ACanvas.TextOut(x + hLine div 2, y, ch);
              fpSuperscript: ACanvas.TextOut(x - hLine div 6, y, ch);
            end;
            dec(y, ACanvas.TextWidth(ch));
          end;
        rtStacked:
          begin
            ACanvas.Font.Orientation := 0;
            w := ACanvas.TextWidth(ch);
            // chars centered around x
            case fntpos of
              fpNormal     : ACanvas.TextOut(x - w div 2, y, ch);
              fpSubscript  : ACanvas.TextOut(x - w div 2, y + hLine div 2, ch);
              fpSuperscript: ACanvas.TextOut(x - w div 2, y - hLine div 6, ch);
            end;
            inc(y, h);
          end;
      end;

      inc(P, UTF8CharacterLength(p));
      if P >= PEnd then break;
    end;
  end;

begin
  if AText = '' then
    exit;

  p := PChar(AText);
  pStartText := p;   // first char of text

  if (Length(ARichTextParams) > 0) then
    iRTP := 0
  else
    iRtp := -1;
  totalHeight := 0;

  if ARotation = rtStacked then
  begin
    Convert_sFont_to_Font(AWorkbook.GetFont(AFontIndex), ACanvas.Font);
    stackPeriod := ACanvas.TextWidth('M') * 2;
  end;

  // Get layout of lines:
  // "lineinfos" collect data on where lines start and end, their width and
  // height, the rich-text parameter index range, and the number of spaces
  // (for text justification)
  repeat
    SetLength(lineInfos, Length(lineInfos)+1);
    with lineInfos[High(lineInfos)] do begin
      pStart := p;
      pEnd := p;
      FirstRtpIndex := iRtp;
      NextRtpIndex := iRtp;
      ScanLine(pEnd, NumSpaces, NextRtpIndex, Width, Height);
      if ARotation = rtStacked then
        totalHeight := totalHeight + stackPeriod
      else
        totalHeight := totalHeight + Height;
      iRtp := NextRtpIndex;
      p := pEnd;
      case p^ of
        ' ': while (p^ <> #0) and (p^ = ' ') do inc(p);
        #13: begin
               inc(p);
               if p^ = #10 then inc(p);
             end;
        #10: inc(p);
      end;
    end;
  until p^ = #0;

  // Draw lines
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
        case AHorAlignment of
          haLeft  : xpos := ARect.Left + stackPeriod div 2;
          haRight : xpos := ARect.Right - totalHeight + stackPeriod div 2;
          haCenter: xpos := (ARect.Left + ARect.Right - totalHeight) div 2;
        end;
      end;
  end;

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
            DrawLine(pStart, pEnd, xpos, ypos, Height, FirstRtpIndex);
            inc(ypos, Height);
          end;
        rt90DegreeClockwiseRotation:
          begin
            case AVertAlignment of
              vaTop    : ypos := ARect.Top;
              vaBottom : ypos := ARect.Bottom - Width;
              vaCenter : ypos := (ARect.Top + ARect.Bottom - Width) div 2;
            end;
            DrawLine(pStart, pEnd, xpos, ypos, Height, FirstRtpIndex);
            dec(xpos, Height);
          end;
        rt90DegreeCounterClockwiseRotation:
          begin
            case AVertAlignment of
              vaTop    : ypos := ARect.Top + Width;
              vaBottom : ypos := ARect.Bottom;
              vaCenter : ypos := (ARect.Top + ARect.Bottom + Width) div 2;
            end;
            DrawLine(pStart, pEnd, xpos, ypos, Height, FirstRtpIndex);
            inc(xpos, Height);
          end;
        rtStacked:
          begin
            case AVertAlignment of
              vaTop    : ypos := ARect.Top;
              vaBottom : ypos := ARect.Bottom - Width;
              vaCenter : ypos := (ARect.Top + ARect.Bottom - Width) div 2;
            end;
            DrawLine(pStart, pEnd, xpos, ypos, Height, FirstRtpIndex);
            inc(xpos, stackPeriod);
          end;
      end;
    end;
  end;
end;

function RichTextWidth(ACanvas: TCanvas; AWorkbook: TsWorkbook; const AText: String;
  AFontIndex: Integer; ARichTextParams: TsRichTextParams): Integer;
var
  s: String;
  p: Integer;
  w, n: Integer;
  rtp, next_rtp: TsRichTextParam;
  fnt, fnt0: TsFont;
begin
  Result := 0;
  if (ACanvas=nil) or (AWorkbook=nil) or (AText = '') then exit;

  fnt0 := AWorkbook.GetFont(AFontIndex);
  Convert_sFont_to_Font(fnt0, ACanvas.Font);

  if Length(ARichTextParams) = 0 then
  begin
    Result := ACanvas.TextWidth(AText);
    if fnt0.Position <> fpNormal then
      Result := Round(Result * SUBSCRIPT_SUPERSCRIPT_FACTOR);
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

end.
