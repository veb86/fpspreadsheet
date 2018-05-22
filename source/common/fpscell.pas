unit fpsCell;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  Classes, SysUtils, fpstypes;

type
  TCellHelper = record helper for TCell
  private
    function GetBackgroundColor: TsColor;
    function GetBiDiMode: TsBiDiMode;
    function GetBorder: TsCellBorders;
    function GetBorderStyle(const ABorder: TsCellBorder): TsCellBorderStyle;
    function GetBorderStyles: TsCellBorderStyles;
    function GetCellFormat: TsCellFormat;
    function GetComment: String;
    function GetFont: TsFont;
    function GetFontIndex: integer;
    function GetHorAlignment: TsHorAlignment;
    function GetHyperlink: TsHyperlink;
    function GetNumberFormat: TsNumberFormat;
    function GetNumberFormatStr: String;
    function GetTextRotation: TsTextRotation;
    function GetUsedFormattingFields: TsUsedFormattingFields;
    function GetVertAlignment: TsVertAlignment;
    function GetWordwrap: Boolean;
    procedure SetBackgroundColor(const AValue: TsColor);
    procedure SetBiDiMode(const AValue: TsBiDiMode);
    procedure SetBorder(const AValue: TsCellBorders);
    procedure SetBorderStyle(const ABorder: TsCellBorder; const AValue: TsCellBorderStyle);
    procedure SetBorderStyles(const AValue: TsCellBorderStyles);
    procedure SetCellFormat(const AValue: TsCellFormat);
    procedure SetComment(const AValue: String);
    procedure SetFontIndex(const AValue: Integer);
    procedure SetHorAlignment(const AValue: TsHorAlignment);
    procedure SetHyperlink(const AValue: TsHyperlink);
    procedure SetNumberFormat(const AValue: TsNumberFormat);
    procedure SetNumberFormatStr(const AValue: String);
    procedure SetTextRotation(const AValue: TsTextRotation);
    procedure SetUsedFormattingFields(const AValue: TsUsedFormattingFields);
    procedure SetVertAlignment(const AValue: TsVertAlignment);
    procedure SetWordwrap(const AValue: Boolean);

  protected
    function GetWorkbook: TsBasicWorkbook; inline;

  public
    property BackgroundColor: TsColor
      read GetBackgroundColor write SetBackgroundColor;
    property BiDiMode: TsBiDiMode
      read GetBiDiMode write SetBiDiMode;
    property Border: TsCellBorders
      read GetBorder write SetBorder;
    property BorderStyle[ABorder: TsCellBorder]: TsCellBorderStyle
      read GetBorderStyle write SetBorderStyle;
    property BorderStyles: TsCellBorderStyles
      read GetBorderStyles write SetBorderStyles;
    property CellFormat: TsCellFormat
      read GetCellFormat write SetCellFormat;
    property Comment: String
      read GetComment write SetComment;
    property Font: TsFont read GetFont;
    property FontIndex: Integer
      read GetFontIndex write SetFontIndex;
    property HorAlignment: TsHorAlignment
      read GetHorAlignment write SetHorAlignment;
    property Hyperlink: TsHyperlink
      read GetHyperlink write SetHyperlink;
    property NumberFormat: TsNumberFormat
      read GetNumberFormat write SetNumberFormat;
    property NumberFormatStr: String
      read GetNumberFormatStr write SetNumberFormatStr;
    property TextRotation: TsTextRotation
      read GetTextRotation write SetTextRotation;
    property UsedFormattingFields: TsUsedFormattingFields
      read GetUsedFormattingFields write SetUsedFormattingFields;
    property VertAlignment: TsVertAlignment
      read GetVertAlignment write SetVertAlignment;
    property Wordwrap: Boolean
      read GetWordwrap write SetWordwrap;
    property Workbook: TsBasicWorkbook read GetWorkbook;
  end;

implementation

uses
  fpspreadsheet;

function TCellHelper.GetBackgroundColor: TsColor;
begin
  Result := (Worksheet as TsWorksheet).ReadBackgroundColor(@self);
end;

function TCellHelper.GetBiDiMode: TsBiDiMode;
begin
  Result := (Worksheet as TsWorksheet).ReadBiDiMode(@self);
end;

function TCellHelper.GetBorder: TsCellBorders;
begin
  Result := (Worksheet as TsWorksheet).ReadCellBorders(@self);
end;

function TCellHelper.GetBorderStyle(const ABorder: TsCellBorder): TsCellBorderStyle;
begin
  Result := (Worksheet as TsWorksheet).ReadCellBorderStyle(@self, ABorder);
end;

function TCellHelper.GetBorderStyles: TsCellBorderStyles;
begin
  Result := (Worksheet as TsWorksheet).ReadCellBorderStyles(@self);
end;

function TCellHelper.GetCellFormat: TsCellFormat;
begin
  Result := (GetWorkbook as TsWorkbook).GetCellFormat(FormatIndex);
end;

function TCellHelper.GetComment: String;
begin
  Result := (Worksheet as TsWorksheet).ReadComment(@self);
end;

function TCellHelper.GetFont: TsFont;
begin
  Result := (Worksheet as TsWorksheet).ReadCellFont(@self);
end;

function TCellHelper.GetFontIndex: Integer;
var
  fmt: PsCellFormat;
begin
  fmt := (Workbook as TsWorkbook).GetPointerToCellFormat(FormatIndex);
  Result := fmt^.FontIndex;
end;

function TCellHelper.GetHorAlignment: TsHorAlignment;
begin
  Result := (Worksheet as TsWorksheet).ReadHorAlignment(@Self);
end;

function TCellHelper.GetHyperlink: TsHyperlink;
begin
  Result := (Worksheet as TsWorksheet).ReadHyperlink(@self);
end;

function TCellHelper.GetNumberFormat: TsNumberFormat;
var
  fmt: PsCellFormat;
begin
  fmt := (Workbook as TsWorkbook).GetPointerToCellFormat(FormatIndex);
  Result := fmt^.NumberFormat;
end;

function TCellHelper.GetNumberFormatStr: String;
var
  fmt: PsCellFormat;
begin
  fmt := (Workbook as TsWorkbook).GetPointerToCellFormat(FormatIndex);
  Result := fmt^.NumberFormatStr;
end;

function TCellHelper.GetTextRotation: TsTextRotation;
begin
  Result := (Worksheet as TsWorksheet).ReadTextRotation(@Self);
end;

function TCellHelper.GetUsedFormattingFields: TsUsedFormattingFields;
begin
  Result := (Worksheet as TsWorksheet).ReadUsedFormatting(@Self);
end;

function TCellHelper.GetVertAlignment: TsVertAlignment;
begin
  Result := (Worksheet as TsWorksheet).ReadVertAlignment(@self);
end;

function TCellHelper.GetWordwrap: Boolean;
begin
  Result := (Worksheet as TsWorksheet).ReadWordwrap(@self);
end;

function TCellHelper.GetWorkbook: TsBasicWorkbook;
begin
  Result := (Worksheet as TsWorksheet).Workbook;
end;

procedure TCellHelper.SetBackgroundColor(const AValue: TsColor);
begin
  (Worksheet as TsWorksheet).WriteBackgroundColor(@self, AValue);
end;

procedure TCellHelper.SetBiDiMode(const AValue: TsBiDiMode);
begin
  (Worksheet as TsWorksheet).WriteBiDiMode(@self, AValue);
end;

procedure TCellHelper.SetBorder(const AValue: TsCellBorders);
begin
  (Worksheet as TsWorksheet).WriteBorders(@self, AValue);
end;

procedure TCellHelper.SetBorderStyle(const ABorder: TsCellBorder;
  const AValue: TsCellBorderStyle);
begin
  (Worksheet as TsWorksheet).WriteBorderStyle(@self, ABorder, AValue);
end;

procedure TCellHelper.SetBorderStyles(const AValue: TsCellBorderStyles);
begin
  (Worksheet as TsWorksheet).WriteBorderStyles(@self, AValue);
end;

procedure TCellHelper.SetCellFormat(const AValue: TsCellFormat);
begin
  (Worksheet as TsWorksheet).WriteCellFormat(@self, AValue);
end;

procedure TCellHelper.SetComment(const AValue: String);
begin
  (Worksheet as TsWorksheet).WriteComment(@self, AValue);
end;

procedure TCellHelper.SetFontIndex(const AValue: Integer);
begin
  (Worksheet as TsWorksheet).WriteFont(@self, AValue);
end;

procedure TCellHelper.SetHorAlignment(const AValue: TsHorAlignment);
begin
  (Worksheet as TsWorksheet).WriteHorAlignment(@self, AValue);
end;

procedure TCellHelper.SetHyperlink(const AValue: TsHyperlink);
begin
  (Worksheet as TsWorksheet).WriteHyperlink(@self, AValue.Target, AValue.Tooltip);
end;

procedure TCellHelper.SetNumberFormat(const AValue: TsNumberFormat);
var
  fmt: TsCellFormat;
begin
  fmt := (Workbook as TsWorkbook).GetCellFormat(FormatIndex);
  fmt.NumberFormat := AValue;
  (Worksheet as TsWorksheet).WriteCellFormat(@self, fmt);
end;

procedure TCellHelper.SetNumberFormatStr(const AValue: String);
var
  fmt: TsCellFormat;
begin
  fmt := (Workbook as TsWorkbook).GetCellFormat(FormatIndex);
  fmt.NumberFormatStr := AValue;
  (Worksheet as TsWorksheet).WriteCellFormat(@self, fmt);
end;

procedure TCellHelper.SetTextRotation(const AValue: TsTextRotation);
begin
  (Worksheet as TsWorksheet).WriteTextRotation(@self, AValue);
end;

procedure TCellHelper.SetUsedFormattingFields(const AValue: TsUsedFormattingFields);
begin
  (Worksheet as TsWorksheet).WriteUsedFormatting(@self, AValue);
end;

procedure TCellHelper.SetVertAlignment(const AValue: TsVertAlignment);
begin
  (Worksheet as TsWorksheet).WriteVertAlignment(@self, AValue);
end;

procedure TCellHelper.SetWordwrap(const AValue: Boolean);
begin
  (Worksheet as TsWorksheet).WriteWordwrap(@self, AValue);
end;


end.

