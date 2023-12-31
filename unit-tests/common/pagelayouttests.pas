{ PageLayout tests
  These unit tests are writing out to and reading back from file.
}

unit pagelayouttests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff8;

type
  { TSpreadWriteReadHyperlinkTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadPageLayoutTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_PageLayout(AFormat: TsSpreadsheetFormat; ANumSheets, ATestMode: Integer);
    procedure TestWriteRead_PageMargins(AFormat: TsSpreadsheetFormat; ANumSheets, AHeaderFooterMode: Integer);
    procedure TestWriteRead_PrintRanges(AFormat: TsSpreadsheetFormat;
      ANumSheets, ANumRanges: Integer; ASpaceInName: Boolean);
    procedure TestWriteRead_RepeatedColRows(AFormat: TsSpreadsheetFormat;
      AFirstCol, ALastCol, AFirstRow, ALastRow: Integer);
      
  published
    { BIFF2 page layout tests }
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_0;
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_1;
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_2;
    procedure TestWriteRead_BIFF2_PageMargins_1sheet_3;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_0;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_1;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_2;
    procedure TestWriteRead_BIFF2_PageMargins_2sheets_3;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_0;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_1;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_2;
    procedure TestWriteRead_BIFF2_PageMargins_3sheets_3;

    procedure TestWriteRead_BIFF2_HeaderFooterRegions_1sheet;
    procedure TestWriteRead_BIFF2_HeaderFooterRegions_2sheets;
    procedure TestWriteRead_BIFF2_HeaderFooterRegions_3sheets;

    procedure TestWriteRead_BIFF2_HeaderFooterSymbols_1sheet;
    procedure TestWriteRead_BIFF2_HeaderFooterSymbols_2sheets;
    procedure TestWriteRead_BIFF2_HeaderFooterSymbols_3sheets;

    procedure TestWriteRead_BIFF2_HeaderFooterFontSymbols_1sheet;
    procedure TestWriteRead_BIFF2_HeaderFooterFontSymbols_2sheets;
    procedure TestWriteRead_BIFF2_HeaderFooterFontSymbols_3sheets;

    // no BIFF2 page orientation tests because this info is not readily available in the file


    { BIFF5 page layout tests }
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_0;
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_1;
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_2;
    procedure TestWriteRead_BIFF5_PageMargins_1sheet_3;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_0;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_1;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_2;
    procedure TestWriteRead_BIFF5_PageMargins_2sheets_3;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_0;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_1;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_2;
    procedure TestWriteRead_BIFF5_PageMargins_3sheets_3;

    procedure TestWriteRead_BIFF5_PageOrientation_1sheet;
    procedure TestWriteRead_BIFF5_PageOrientation_2sheets;
    procedure TestWriteRead_BIFF5_PageOrientation_3sheets;

    procedure TestWriteRead_BIFF5_PaperSize_1sheet;
    procedure TestWriteRead_BIFF5_PaperSize_2sheets;
    procedure TestWriteRead_BIFF5_PaperSize_3sheets;

    procedure TestWriteRead_BIFF5_ScalingFactor_1sheet;
    procedure TestWriteRead_BIFF5_ScalingFactor_2sheets;
    procedure TestWriteRead_BIFF5_ScalingFactor_3sheets;

    procedure TestWriteRead_BIFF5_WidthToPages_1sheet;
    procedure TestWriteRead_BIFF5_WidthToPages_2sheets;
    procedure TestWriteRead_BIFF5_WidthToPages_3sheets;

    procedure TestWriteRead_BIFF5_HeightToPages_1sheet;
    procedure TestWriteRead_BIFF5_HeightToPages_2sheets;
    procedure TestWriteRead_BIFF5_HeightToPages_3sheets;

    procedure TestWriteRead_BIFF5_PageNumber_1sheet;
    procedure TestWriteRead_BIFF5_PageNumber_2sheets;
    procedure TestWriteRead_BIFF5_PageNumber_3sheets;

    procedure TestWriteRead_BIFF5_HeaderFooterRegions_1sheet;
    procedure TestWriteRead_BIFF5_HeaderFooterRegions_2sheets;
    procedure TestWriteRead_BIFF5_HeaderFooterRegions_3sheets;

    procedure TestWriteRead_BIFF5_HeaderFooterSymbols_1sheet;
    procedure TestWriteRead_BIFF5_HeaderFooterSymbols_2sheets;
    procedure TestWriteRead_BIFF5_HeaderFooterSymbols_3sheets;

    procedure TestWriteRead_BIFF5_HeaderFooterFontSymbols_1sheet;
    procedure TestWriteRead_BIFF5_HeaderFooterFontSymbols_2sheets;
    procedure TestWriteRead_BIFF5_HeaderFooterFontSymbols_3sheets;

    procedure TestWriteRead_BIFF5_HeaderFooterFontColor_1sheet;
    procedure TestWriteRead_BIFF5_HeaderFooterFontColor_2sheets;
    procedure TestWriteRead_BIFF5_HeaderFooterFontColor_3sheets;

    procedure TestWriteRead_BIFF5_PrintRanges_1sheet_1Range_NoSpace;
    procedure TestWriteRead_BIFF5_PrintRanges_1sheet_2Ranges_NoSpace;
    procedure TestWriteRead_BIFF5_PrintRanges_2sheet_1Range_NoSpace;
    procedure TestWriteRead_BIFF5_PrintRanges_2sheet_2Ranges_NoSpace;

    procedure TestWriteRead_BIFF5_PrintRanges_1sheet_1Range_Space;
    procedure TestWriteRead_BIFF5_PrintRanges_1sheet_2Ranges_Space;
    procedure TestWriteRead_BIFF5_PrintRanges_2sheet_1Range_Space;
    procedure TestWriteRead_BIFF5_PrintRanges_2sheet_2Ranges_Space;

    procedure TestWriteRead_BIFF5_RepeatedRow_0;
    procedure TestWriteRead_BIFF5_RepeatedRows_0_1;
    procedure TestWriteRead_BIFF5_RepeatedRows_1_3;
    procedure TestWriteRead_BIFF5_RepeatedCol_0;
    procedure TestWriteRead_BIFF5_RepeatedCols_0_1;
    procedure TestWriteRead_BIFF5_RepeatedCols_1_3;
    procedure TestWriteRead_BIFF5_RepeatedCol_0_Row_0;
    procedure TestWriteRead_BIFF5_RepeatedCols_0_1_Rows_0_1;

    { BIFF8 page layout tests }
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_0;
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_1;
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_2;
    procedure TestWriteRead_BIFF8_PageMargins_1sheet_3;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_0;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_1;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_2;
    procedure TestWriteRead_BIFF8_PageMargins_2sheets_3;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_0;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_1;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_2;
    procedure TestWriteRead_BIFF8_PageMargins_3sheets_3;

    procedure TestWriteRead_BIFF8_PageOrientation_1sheet;
    procedure TestWriteRead_BIFF8_PageOrientation_2sheets;
    procedure TestWriteRead_BIFF8_PageOrientation_3sheets;

    procedure TestWriteRead_BIFF8_PaperSize_1sheet;
    procedure TestWriteRead_BIFF8_PaperSize_2sheets;
    procedure TestWriteRead_BIFF8_PaperSize_3sheets;

    procedure TestWriteRead_BIFF8_ScalingFactor_1sheet;
    procedure TestWriteRead_BIFF8_ScalingFactor_2sheets;
    procedure TestWriteRead_BIFF8_ScalingFactor_3sheets;

    procedure TestWriteRead_BIFF8_WidthToPages_1sheet;
    procedure TestWriteRead_BIFF8_WidthToPages_2sheets;
    procedure TestWriteRead_BIFF8_WidthToPages_3sheets;

    procedure TestWriteRead_BIFF8_HeightToPages_1sheet;
    procedure TestWriteRead_BIFF8_HeightToPages_2sheets;
    procedure TestWriteRead_BIFF8_HeightToPages_3sheets;

    procedure TestWriteRead_BIFF8_PageNumber_1sheet;
    procedure TestWriteRead_BIFF8_PageNumber_2sheets;
    procedure TestWriteRead_BIFF8_PageNumber_3sheets;

    procedure TestWriteRead_BIFF8_HeaderFooterRegions_1sheet;
    procedure TestWriteRead_BIFF8_HeaderFooterRegions_2sheets;
    procedure TestWriteRead_BIFF8_HeaderFooterRegions_3sheets;

    procedure TestWriteRead_BIFF8_HeaderFooterSymbols_1sheet;
    procedure TestWriteRead_BIFF8_HeaderFooterSymbols_2sheets;
    procedure TestWriteRead_BIFF8_HeaderFooterSymbols_3sheets;

    procedure TestWriteRead_BIFF8_HeaderFooterFontSymbols_1sheet;
    procedure TestWriteRead_BIFF8_HeaderFooterFontSymbols_2sheets;
    procedure TestWriteRead_BIFF8_HeaderFooterFontSymbols_3sheets;

    procedure TestWriteRead_BIFF8_HeaderFooterFontColor_1sheet;
    procedure TestWriteRead_BIFF8_HeaderFooterFontColor_2sheets;
    procedure TestWriteRead_BIFF8_HeaderFooterFontColor_3sheets;

    procedure TestWriteRead_BIFF8_PrintRanges_1sheet_1Range_NoSpace;
    procedure TestWriteRead_BIFF8_PrintRanges_1sheet_2Ranges_NoSpace;
    procedure TestWriteRead_BIFF8_PrintRanges_2sheet_1Range_NoSpace;
    procedure TestWriteRead_BIFF8_PrintRanges_2sheet_2Ranges_NoSpace;

    procedure TestWriteRead_BIFF8_PrintRanges_1sheet_1Range_Space;
    procedure TestWriteRead_BIFF8_PrintRanges_1sheet_2Ranges_Space;
    procedure TestWriteRead_BIFF8_PrintRanges_2sheet_1Range_Space;
    procedure TestWriteRead_BIFF8_PrintRanges_2sheet_2Ranges_Space;

    procedure TestWriteRead_BIFF8_RepeatedRow_0;
    procedure TestWriteRead_BIFF8_RepeatedRows_0_1;
    procedure TestWriteRead_BIFF8_RepeatedRows_1_3;
    procedure TestWriteRead_BIFF8_RepeatedCol_0;
    procedure TestWriteRead_BIFF8_RepeatedCols_0_1;
    procedure TestWriteRead_BIFF8_RepeatedCols_1_3;
    procedure TestWriteRead_BIFF8_RepeatedCol_0_Row_0;
    procedure TestWriteRead_BIFF8_RepeatedCols_0_1_Rows_0_1;

    { OOXML page layout tests }
    procedure TestWriteRead_OOXML_PageMargins_1sheet_0;
    procedure TestWriteRead_OOXML_PageMargins_1sheet_1;
    procedure TestWriteRead_OOXML_PageMargins_1sheet_2;
    procedure TestWriteRead_OOXML_PageMargins_1sheet_3;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_0;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_1;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_2;
    procedure TestWriteRead_OOXML_PageMargins_2sheets_3;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_0;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_1;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_2;
    procedure TestWriteRead_OOXML_PageMargins_3sheets_3;

    procedure TestWriteRead_OOXML_PageOrientation_1sheet;
    procedure TestWriteRead_OOXML_PageOrientation_2sheets;
    procedure TestWriteRead_OOXML_PageOrientation_3sheets;

    procedure TestWriteRead_OOXML_PaperSize_1sheet;
    procedure TestWriteRead_OOXML_PaperSize_2sheets;
    procedure TestWriteRead_OOXML_PaperSize_3sheets;

    procedure TestWriteRead_OOXML_ScalingFactor_1sheet;
    procedure TestWriteRead_OOXML_ScalingFactor_2sheets;
    procedure TestWriteRead_OOXML_ScalingFactor_3sheets;

    procedure TestWriteRead_OOXML_WidthToPages_1sheet;
    procedure TestWriteRead_OOXML_WidthToPages_2sheets;
    procedure TestWriteRead_OOXML_WidthToPages_3sheets;

    procedure TestWriteRead_OOXML_HeightToPages_1sheet;
    procedure TestWriteRead_OOXML_HeightToPages_2sheets;
    procedure TestWriteRead_OOXML_HeightToPages_3sheets;

    procedure TestWriteRead_OOXML_PageNumber_1sheet;
    procedure TestWriteRead_OOXML_PageNumber_2sheets;
    procedure TestWriteRead_OOXML_PageNumber_3sheets;

    procedure TestWriteRead_OOXML_HeaderFooterRegions_1sheet;
    procedure TestWriteRead_OOXML_HeaderFooterRegions_2sheets;
    procedure TestWriteRead_OOXML_HeaderFooterRegions_3sheets;

    procedure TestWriteRead_OOXML_HeaderFooterSymbols_1sheet;
    procedure TestWriteRead_OOXML_HeaderFooterSymbols_2sheets;
    procedure TestWriteRead_OOXML_HeaderFooterSymbols_3sheets;

    procedure TestWriteRead_OOXML_HeaderFooterFontSymbols_1sheet;
    procedure TestWriteRead_OOXML_HeaderFooterFontSymbols_2sheets;
    procedure TestWriteRead_OOXML_HeaderFooterFontSymbols_3sheets;

    procedure TestWriteRead_OOXML_HeaderFooterFontColor_1sheet;
    procedure TestWriteRead_OOXML_HeaderFooterFontColor_2sheets;
    procedure TestWriteRead_OOXML_HeaderFooterFontColor_3sheets;

    procedure TestWriteRead_OOXML_PrintRanges_1sheet_1Range_NoSpace;
    procedure TestWriteRead_OOXML_PrintRanges_1sheet_2Ranges_NoSpace;
    procedure TestWriteRead_OOXML_PrintRanges_2sheet_1Range_NoSpace;
    procedure TestWriteRead_OOXML_PrintRanges_2sheet_2Ranges_NoSpace;

    procedure TestWriteRead_OOXML_PrintRanges_1sheet_1Range_Space;
    procedure TestWriteRead_OOXML_PrintRanges_1sheet_2Ranges_Space;
    procedure TestWriteRead_OOXML_PrintRanges_2sheet_1Range_Space;
    procedure TestWriteRead_OOXML_PrintRanges_2sheet_2Ranges_Space;

    procedure TestWriteRead_OOXML_RepeatedRow_0;
    procedure TestWriteRead_OOXML_RepeatedRows_0_1;
    procedure TestWriteRead_OOXML_RepeatedRows_1_3;
    procedure TestWriteRead_OOXML_RepeatedCol_0;
    procedure TestWriteRead_OOXML_RepeatedCols_0_1;
    procedure TestWriteRead_OOXML_RepeatedCols_1_3;
    procedure TestWriteRead_OOXML_RepeatedCol_0_Row_0;
    procedure TestWriteRead_OOXML_RepeatedCols_0_1_Rows_0_1;

    { Excel2003/XML page layout tests }
    procedure TestWriteRead_XML_PageMargins_1sheet_0;
    procedure TestWriteRead_XML_PageMargins_1sheet_1;
    procedure TestWriteRead_XML_PageMargins_1sheet_2;
    procedure TestWriteRead_XML_PageMargins_1sheet_3;
    procedure TestWriteRead_XML_PageMargins_2sheets_0;
    procedure TestWriteRead_XML_PageMargins_2sheets_1;
    procedure TestWriteRead_XML_PageMargins_2sheets_2;
    procedure TestWriteRead_XML_PageMargins_2sheets_3;
    procedure TestWriteRead_XML_PageMargins_3sheets_0;
    procedure TestWriteRead_XML_PageMargins_3sheets_1;
    procedure TestWriteRead_XML_PageMargins_3sheets_2;
    procedure TestWriteRead_XML_PageMargins_3sheets_3;

    procedure TestWriteRead_XML_PageOrientation_1sheet;
    procedure TestWriteRead_XML_PageOrientation_2sheets;
    procedure TestWriteRead_XML_PageOrientation_3sheets;

    procedure TestWriteRead_XML_PaperSize_1sheet;
    procedure TestWriteRead_XML_PaperSize_2sheets;
    procedure TestWriteRead_XML_PaperSize_3sheets;

    procedure TestWriteRead_XML_ScalingFactor_1sheet;
    procedure TestWriteRead_XML_ScalingFactor_2sheets;
    procedure TestWriteRead_XML_ScalingFactor_3sheets;

    procedure TestWriteRead_XML_WidthToPages_1sheet;
    procedure TestWriteRead_XML_WidthToPages_2sheets;
    procedure TestWriteRead_XML_WidthToPages_3sheets;

    procedure TestWriteRead_XML_HeightToPages_1sheet;
    procedure TestWriteRead_XML_HeightToPages_2sheets;
    procedure TestWriteRead_XML_HeightToPages_3sheets;

    procedure TestWriteRead_XML_PageNumber_1sheet;
    procedure TestWriteRead_XML_PageNumber_2sheets;
    procedure TestWriteRead_XML_PageNumber_3sheets;

    procedure TestWriteRead_XML_HeaderFooterRegions_1sheet;
    procedure TestWriteRead_XML_HeaderFooterRegions_2sheets;
    procedure TestWriteRead_XML_HeaderFooterRegions_3sheets;

    procedure TestWriteRead_XML_HeaderFooterSymbols_1sheet;
    procedure TestWriteRead_XML_HeaderFooterSymbols_2sheets;
    procedure TestWriteRead_XML_HeaderFooterSymbols_3sheets;

    procedure TestWriteRead_XML_HeaderFooterFontSymbols_1sheet;
    procedure TestWriteRead_XML_HeaderFooterFontSymbols_2sheets;
    procedure TestWriteRead_XML_HeaderFooterFontSymbols_3sheets;

    procedure TestWriteRead_XML_HeaderFooterFontColor_1sheet;
    procedure TestWriteRead_XML_HeaderFooterFontColor_2sheets;
    procedure TestWriteRead_XML_HeaderFooterFontColor_3sheets;

    procedure TestWriteRead_XML_PrintRanges_1sheet_1Range_NoSpace;
    procedure TestWriteRead_XML_PrintRanges_1sheet_2Ranges_NoSpace;
    procedure TestWriteRead_XML_PrintRanges_2sheet_1Range_NoSpace;
    procedure TestWriteRead_XML_PrintRanges_2sheet_2Ranges_NoSpace;

    procedure TestWriteRead_XML_PrintRanges_1sheet_1Range_Space;
    procedure TestWriteRead_XML_PrintRanges_1sheet_2Ranges_Space;
    procedure TestWriteRead_XML_PrintRanges_2sheet_1Range_Space;
    procedure TestWriteRead_XML_PrintRanges_2sheet_2Ranges_Space;

    procedure TestWriteRead_XML_RepeatedRow_0;
    procedure TestWriteRead_XML_RepeatedRows_0_1;
    procedure TestWriteRead_XML_RepeatedRows_1_3;
    procedure TestWriteRead_XML_RepeatedCol_0;
    procedure TestWriteRead_XML_RepeatedCols_0_1;
    procedure TestWriteRead_XML_RepeatedCols_1_3;
    procedure TestWriteRead_XML_RepeatedCol_0_Row_0;
    procedure TestWriteRead_XML_RepeatedCols_0_1_Rows_0_1;

    { OpenDocument page layout tests }
    procedure TestWriteRead_ODS_PageMargins_1sheet_0;
    procedure TestWriteRead_ODS_PageMargins_1sheet_1;
    procedure TestWriteRead_ODS_PageMargins_1sheet_2;
    procedure TestWriteRead_ODS_PageMargins_1sheet_3;
    procedure TestWriteRead_ODS_PageMargins_2sheets_0;
    procedure TestWriteRead_ODS_PageMargins_2sheets_1;
    procedure TestWriteRead_ODS_PageMargins_2sheets_2;
    procedure TestWriteRead_ODS_PageMargins_2sheets_3;
    procedure TestWriteRead_ODS_PageMargins_3sheets_0;
    procedure TestWriteRead_ODS_PageMargins_3sheets_1;
    procedure TestWriteRead_ODS_PageMargins_3sheets_2;
    procedure TestWriteRead_ODS_PageMargins_3sheets_3;

    procedure TestWriteRead_ODS_PageOrientation_1sheet;
    procedure TestWriteRead_ODS_PageOrientation_2sheets;
    procedure TestWriteRead_ODS_PageOrientation_3sheets;

    procedure TestWriteRead_ODS_PaperSize_1sheet;
    procedure TestWriteRead_ODS_PaperSize_2sheets;
    procedure TestWriteRead_ODS_PaperSize_3sheets;

    procedure TestWriteRead_ODS_ScalingFactor_1sheet;
    procedure TestWriteRead_ODS_ScalingFactor_2sheets;
    procedure TestWriteRead_ODS_ScalingFactor_3sheets;

    procedure TestWriteRead_ODS_WidthToPages_1sheet;
    procedure TestWriteRead_ODS_WidthToPages_2sheets;
    procedure TestWriteRead_ODS_WidthToPages_3sheets;

    procedure TestWriteRead_ODS_HeightToPages_1sheet;
    procedure TestWriteRead_ODS_HeightToPages_2sheets;
    procedure TestWriteRead_ODS_HeightToPages_3sheets;

    procedure TestWriteRead_ODS_PageNumber_1sheet;
    procedure TestWriteRead_ODS_PageNumber_2sheets;
    procedure TestWriteRead_ODS_PageNumber_3sheets;

    procedure TestWriteRead_ODS_HeaderFooterRegions_1sheet;
    procedure TestWriteRead_ODS_HeaderFooterRegions_2sheets;
    procedure TestWriteRead_ODS_HeaderFooterRegions_3sheets;

    procedure TestWriteRead_ODS_HeaderFooterSymbols_1sheet;
    procedure TestWriteRead_ODS_HeaderFooterSymbols_2sheets;
    procedure TestWriteRead_ODS_HeaderFooterSymbols_3sheets;

    procedure TestWriteRead_ODS_HeaderFooterFontSymbols_1sheet;
    procedure TestWriteRead_ODS_HeaderFooterFontSymbols_2sheets;
    procedure TestWriteRead_ODS_HeaderFooterFontSymbols_3sheets;

    procedure TestWriteRead_ODS_HeaderFooterFontColor_1sheet;
    procedure TestWriteRead_ODS_HeaderFooterFontColor_2sheets;
    procedure TestWriteRead_ODS_HeaderFooterFontColor_3sheets;

    procedure TestWriteRead_ODS_PrintRanges_1sheet_1Range_NoSpace;
    procedure TestWriteRead_ODS_PrintRanges_1sheet_2Ranges_NoSpace;
    procedure TestWriteRead_ODS_PrintRanges_2sheet_1Range_NoSpace;
    procedure TestWriteRead_ODS_PrintRanges_2sheet_2Ranges_NoSpace;

    procedure TestWriteRead_ODS_PrintRanges_1sheet_1Range_Space;
    procedure TestWriteRead_ODS_PrintRanges_1sheet_2Ranges_Space;
    procedure TestWriteRead_ODS_PrintRanges_2sheet_1Range_Space;
    procedure TestWriteRead_ODS_PrintRanges_2sheet_2Ranges_Space;

    procedure TestWriteRead_ODS_RepeatedRow_0;
    procedure TestWriteRead_ODS_RepeatedRows_0_1;
    procedure TestWriteRead_ODS_RepeatedRows_1_3;
    procedure TestWriteRead_ODS_RepeatedCol_0;
    procedure TestWriteRead_ODS_RepeatedCols_0_1;
    procedure TestWriteRead_ODS_RepeatedCols_1_3;
    procedure TestWriteRead_ODS_RepeatedCol_0_Row_0;
    procedure TestWriteRead_ODS_RepeatedCols_0_1_Rows_0_1;
  end;

implementation

uses
  typinfo, contnrs, strutils,
  fpsutils, fpsHeaderFooterParser, fpsPageLayout;
//  uriparser, lazfileutils, fpsutils;

const
  PageLayoutSheet = 'PageLayout';

const
  SollRanges: Array[1..2] of TsCellRange = (
    (Row1: 0; Col1: 0; Row2:10; Col2:20),
    (Row1:20; Col1:30; Row2:25; Col2:40)
  );


{ TSpreadWriteReadPageLayoutTests }

procedure TSpreadWriteReadPageLayoutTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadPageLayoutTests.TearDown;
begin
  inherited TearDown;
end;

{ AHeaderFooterMode = 0 ... no header, no footer
                      1 ... header, no footer
                      2 ... no header, footer
	                  3 ... header, footer }
procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageMargins(
  AFormat: TsSpreadsheetFormat; ANumSheets, AHeaderFooterMode: Integer);
const
  EPS = 1e-6;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col, p: Integer;
  sollPageLayout, actualPageLayout: TsPageLayout;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;

  sollPageLayout := TsPageLayout.Create(nil);
  try
    with SollPageLayout do
    begin
      TopMargin := 20;
      BottomMargin := 30;
      LeftMargin := 21;
      RightMargin := 22;
      HeaderMargin := 10;
      FooterMargin := 11;
      case AHeaderFooterMode of
        0: ;  // header and footer already are empty strings
        1: Headers[HEADER_FOOTER_INDEX_ALL] := 'Test header';
        2: Footers[HEADER_FOOTER_INDEX_ALL] := 'Test footer';
        3: begin
             Headers[HEADER_FOOTER_INDEX_ALL] := 'Test header';
  	       Footers[HEADER_FOOTER_INDEX_ALL] := 'Test footer';
           end;
      end;
    end;
  
    MyWorkbook := TsWorkbook.Create;
    try
      col := 0;
      for p := 1 to ANumSheets do
      begin
        MyWorkSheet:= MyWorkBook.AddWorksheet(PageLayoutSheet+IntToStr(p));
        for row := 0 to 9 do
          Myworksheet.WriteNumber(row, 0, row+col*100+p*10000 );
        MyWorksheet.PageLayout.Assign(SollPageLayout);
      end;
      MyWorkBook.WriteToFile(TempFile, AFormat, true);
    finally
      MyWorkbook.Free;
    end;

    // Open the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.ReadFromFile(TempFile, AFormat);
      for p := 0 to MyWorkbook.GetWorksheetCount-1 do
      begin
        MyWorksheet := MyWorkBook.GetWorksheetByIndex(p);
        if MyWorksheet=nil then
          fail('Error in test code. Failed to get worksheet by index');
	
        actualPageLayout := MyWorksheet.PageLayout;
        CheckEquals(sollPageLayout.TopMargin, actualPageLayout.TopMargin, EPS, 'Top margin mismatch, sheet "'+MyWorksheet.Name+'"');
        CheckEquals(sollPageLayout.BottomMargin, actualPageLayout.Bottommargin, EPS, 'Bottom margin mismatch, sheet "'+MyWorksheet.Name+'"');
        CheckEquals(sollPageLayout.LeftMargin, actualPageLayout.LeftMargin, EPS, 'Left margin mismatch, sheet "'+MyWorksheet.Name+'"');
        CheckEquals(sollPageLayout.RightMargin, actualPageLayout.RightMargin, EPS, 'Right margin mismatch, sheet "'+MyWorksheet.Name+'"');
        if (AFormat <> sfExcel2) then  // No header/footer margin in BIFF2
        begin
          if AHeaderFooterMode in [1, 3] then
            CheckEquals(sollPageLayout.HeaderMargin, actualPageLayout.HeaderMargin, EPS, 'Header margin mismatch, sheet "'+MyWorksheet.Name+'"');
          if AHeaderFooterMode in [2, 3] then
            CheckEquals(sollPageLayout.FooterMargin, actualPageLayout.FooterMargin, EPS, 'Footer margin mismatch, sheet "'+MyWorksheet.Name+'"');
        end;
      end;

    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

  finally
    SollPageLayout.Free;
  end;
end;

{ ------------------------------------------------------------------------------
 Main page layout test: it writes a file with a specific page layout and reads it
 back. The written pagelayout ("SollLayout") must match the read pagelayout.

 ATestMode:
   0 - Landscape page orientation for sheets 0 und 2, sheet 1 is portrait
   1 - Paper size: sheet 1 "Letter" (8.5" x 11"), sheets 0 and 2 "A5" (148 mm x 210 mm)
   2 - Scaling factor: sheet 1 50%, sheet 2 200%, sheet 3 100%
   3 - Scale n pages to width: sheet 1 n=2, sheet 2 n=3, sheet 3 n=1
   4 - Scale n pages to height: sheet 1 n=2, sheet 2 n=3, sheet 3 n=1
   5 - First page number: sheet 1 - 3, sheet 2 - automatic, sheet 3 - 1
   6 - Header/footer region test: sheet 1 - header only, sheet 2 - footer only, sheet 3 - both
-------------------------------------------------------------------------------}
procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PageLayout(
  AFormat: TsSpreadsheetFormat; ANumSheets, ATestMode: Integer);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col, p: Integer;
  sollPageLayout: Array of TsPageLayout = nil;
  actualPageLayout: TsPageLayout;
  TempFile: string; //write xls/xml to this file and read back from it

  function SameParsedHeaderFooter(AText1, AText2: String;
    AWorkbook: TsWorkbook): Boolean;
  var
    parser1, parser2: TsHeaderFooterParser;
    list1, list2: TObjectList;
    s: TsHeaderFooterSectionIndex;
    el: Integer;
    defFnt: TsHeaderFooterFont;
  begin
    Result := false;
    list1 := TObjectList.Create;
    list2 := TObjectList.Create;
    defFnt := TsHeaderFooterFont.Create(AWorkbook.GetDefaultFont);
    try
      parser1 := TsHeaderFooterParser.Create(AText1, list1, defFnt);
      parser2 := TsHeaderFooterParser.Create(AText2, list2, defFnt);
      try
        for s := Low(TsHeaderFooterSectionIndex) to High(TsHeaderFooterSectionIndex) do
        begin
          if Length(parser1.Sections[s]) <> Length(parser2.Sections[s]) then
            exit;
          for el := 0 to Length(parser1.Sections[s])-1 do
          begin
            if parser1.Sections[s][el].Token <> parser2.Sections[s][el].Token then
              exit;
            if parser1.Sections[s][el].TextValue <> parser2.Sections[s][el].TextValue then
              exit;
            if parser1.Sections[s][el].FontIndex <> parser2.Sections[s][el].FontIndex then
              exit;
          end;
        end;
        Result := true;
      finally
        parser1.Free;
        parser2.Free;
      end;
    finally;
      defFnt.Free;
      list1.Free;
      list2.Free;
    end;
  end;

begin
  TempFile := GetTempFileName;

  SetLength(SollPageLayout, ANumSheets);
  for p:=0 to High(SollPageLayout) do
  begin
    sollPageLayout[p] := TsPageLayout.Create(nil);
    with SollPageLayout[p] do
    begin
      case ATestMode of
        0: // Page orientation test: sheets 0 and 2 are portrait, sheet 1 is landscape
           if p <> 1 then Orientation := spoLandscape;
        1: // Paper size test: sheets 0 and 2 are A5, sheet 1 is LETTER
           if odd(p) then
           begin
             PageWidth := 8.5*2.54; PageHeight := 11*2.54;
           end else
           begin
             PageWidth := 148; PageHeight := 210;
           end;
        2: // Scaling factor: sheet 1 50%, sheet 2 200%, sheet 3 100%
           begin
             if p = 0 then ScalingFactor := 50 else
             if p = 1 then ScalingFactor := 200;
           end;
        3: // Scale width to n pages
           begin
             case p of
               0: FitWidthToPages := 2;
               1: FitWidthToPages := 3;
               2: FitWidthToPages := 1;
             end;
           end;
        4: // Scale height to n pages
           begin
             case p of
               0: FitHeightToPages := 2;
               1: FitHeightToPages := 3;
               2: FitHeightToPages := 1;
             end;
           end;
        5: // Page number of first pge
           begin
             case p of
               0: StartPageNumber := 3;
               1: Options := Options - [poUseStartPageNumber];
               2: StartPageNumber := 1;
             end;
             Headers[HEADER_FOOTER_INDEX_ALL] := '&LPage &P of &N';
           end;
        6: // Header/footer region test
           case p of
             0: Headers[HEADER_FOOTER_INDEX_ALL] := '&LLeft header&CCenter header&RRight header';
             1: Footers[HEADER_FOOTER_INDEX_ALL] := '&LLeft foorer&CCenter footer&RRight footer';
             2: begin
                  Headers[HEADER_FOOTER_INDEX_ALL] := '&LLeft header&CCenter header&RRight header';
                  Footers[HEADER_FOOTER_INDEX_ALL] := '&LLeft foorer&CCenter footer&RRight footer';
                end;
           end;
        7: // Header/footer symbol test
           case p of
             0: Headers[HEADER_FOOTER_INDEX_ALL] := '&LPage &P / Page count &N&CDate &D - Time &T&RFile &Z&F';
             1: Footers[HEADER_FOOTER_INDEX_ALL] := '&LSheet "&A"&C100&&';
             2: begin
                  Headers[HEADER_FOOTER_INDEX_ALL] := '&LPage &P of &N&C&D &T&R&Z&F';
                  Footers[HEADER_FOOTER_INDEX_ALL] := '&LSheet "&A"&C100&&';
                end;
           end;
        8: // Header/footer font symbol test
           begin
             Headers[HEADER_FOOTER_INDEX_ALL] :=
               '&LH'+
                 '&Y2&YO cm&X2'+
               '&C'+
                 '&"Times New Roman"&18This is big'+
               '&R'+
                 'This is &Bbold&B,'+ LineEnding+'&Iitalic&I,'+LineEnding+
                 '&Uunderlined&U,'+LineEnding+'&Edouble underlined&E,'+
                 '&Sstriked-out&S,'+LineEnding+'&Ooutlined&O,'+LineEnding+
                 '&Hshadow';
             Footers[HEADER_FOOTER_INDEX_ALL] :=
               '&L&"Arial"&8Arial small'+
               '&C&"Courier new"&32Courier big'+
               '&R&"Times New Roman"&10Times standard';
             case p of
               0: Footers[HEADER_FOOTER_INDEX_ALL] := '';
               1: Headers[HEADER_FOOTER_INDEX_ALL] := '';
             end;
           end;
        9: // Header/footer font color test
           begin
             Headers[HEADER_FOOTER_INDEX_ALL] :=
               '&L&KFF0000This is red'+
               '&C&K00FF00This is green'+
               '&R&K0000FFThis is blue';
             Footers[HEADER_FOOTER_INDEX_ALL] :=
               '&LThis is &"Times New Roman"&KFF0000red&K000000, &K00FF00green&K000000, &K0000FFblue&K000000.';
             case p of
               0: Footers[HEADER_FOOTER_INDEX_ALL] := '';
               1: Headers[HEADER_FOOTER_INDEX_ALL] := '';
             end;
           end;

      end;
    end;
  end;

  MyWorkbook := TsWorkbook.Create;
  try
    for p := 0 to ANumSheets-1 do
    begin
      MyWorkSheet:= MyWorkBook.AddWorksheet(PageLayoutSheet+IntToStr(p+1));
      for row := 0 to 99 do
        for col := 0 to 29 do
          Myworksheet.WriteNumber(row, col, (row+1)+(col+1)*100+(p+1)*10000 );
      MyWorksheet.PageLayout.Assign(SollPageLayout[p]);
    end;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    for p := 0 to MyWorkbook.GetWorksheetCount-1 do
    begin
      MyWorksheet := MyWorkBook.GetWorksheetByIndex(p);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get worksheet by index');

      actualPageLayout := MyWorksheet.PageLayout;
      case ATestMode of
        0: // Page orientation test
          CheckEquals(GetEnumName(TypeInfo(TsPageOrientation), ord(sollPageLayout[p].Orientation)),
            GetEnumName(TypeInfo(TsPageOrientation), ord(actualPageLayout.Orientation)),
           'Page orientation mismatch, sheet "'+MyWorksheet.Name+'"'
          );
        1: // Paper size test
          begin
            CheckEquals(sollPagelayout[p].PageHeight, actualPageLayout.PageHeight, 0.1,
              'Page height mismatch, sheet "' + MyWorksheet.Name + '"');
            CheckEquals(sollPageLayout[p].PageWidth, actualPageLayout.PageWidth, 0.1,
              'Page width mismatch, sheet "' + MyWorksheet.name + '"');
          end;
        2: // Scaling factor
          begin
            CheckEquals(poFitPages in sollPageLayout[p].Options, poFitPages in actualPageLayout.Options,
              '"poFitPages" option mismatch, sheet "' + MyWorksheet.name + '"');
            CheckEquals(sollPageLayout[p].ScalingFactor, actualPageLayout.ScalingFactor,
              'Scaling factor mismatch, sheet "' + MyWorksheet.Name + '"');
          end;
        3: // Fit width to pages
          begin
            CheckEquals(poFitPages in sollPageLayout[p].Options, poFitPages in actualPageLayout.Options,
              '"poFitPages" option mismatch, sheet "' + MyWorksheet.name + '"');
            CheckEquals(sollPageLayout[p].FitWidthToPages, actualPageLayout.FitWidthToPages,
              'FitWidthToPages mismatch, sheet "' + MyWorksheet.Name + '"');
          end;
        4: // Fit height to pages
          begin
            CheckEquals(poFitPages in sollPageLayout[p].Options, poFitPages in actualPageLayout.Options,
              '"poFitPages" option mismatch, sheet "' + MyWorksheet.name + '"');
            CheckEquals(sollPageLayout[p].FitHeightToPages, actualPageLayout.FitHeightToPages,
              'FitWidthToPages mismatch, sheet "' + MyWorksheet.Name + '"');
          end;
        5: // Start page number
          begin
            CheckEquals(poUseStartPageNumber in sollPageLayout[p].Options, poUseStartPageNumber in actualPageLayout.Options,
              '"poUseStartPageNumber" option mismatch, sheet "' + MyWorksheet.name + '"');
            CheckEquals(sollPageLayout[p].StartPageNumber, actualPageLayout.StartPageNumber,
              'StartPageNumber value mismatch, sheet "' + MyWorksheet.Name + '"');
          end;
        6, 7, 8, 9: // Header/footer tests
          begin
            if (sollPageLayout[p].Headers[1] <> actualPageLayout.Headers[1]) and
              not SameParsedHeaderFooter(sollPagelayout[p].Headers[1], actualPageLayout.Headers[1], MyWorkbook)
            then
              CheckEquals(sollPageLayout[p].Headers[1], actualPageLayout.Headers[1],
                'Header value mismatch, sheet "' + MyWorksheet.Name + '"');
            if (sollPageLayout[p].Footers[1] <> actualPageLayout.Footers[1]) and
              not SameParsedHeaderFooter(sollPagelayout[p].Footers[1], actualPageLayout.Footers[1], MyWorkbook)
            then
              CheckEquals(sollPageLayout[p].Footers[1], actualPageLayout.Footers[1],
                'Footer value mismatch, sheet "' + MyWorksheet.Name + '"');
          end;
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;

  for p:=0 to High(SollPageLayout) do
    sollPageLayout[p].Free;
end;

{
soll:
'&LH&Y2&YO cm&X2&X&C&"Times New Roman"&18This is big&RThis is &Bbold&B,'#13#10'&Iitalic&I,'#13#10'&Uunderlined&U,'#13#10'&Edouble underlined&E,&Sstriked-out&S,'#13#10'&Ooutlined&O,'#13#10'&Hshadow&H'

actual:
'&LH&Y2&YO cm&X2  &C&"Times New Roman"&18This is big&RThis is &Bbold&B,'#13#10'&Iitalic&I,'#13#10'&Uunderlined&U,'#13#10'&Edouble underlined&E,&Sstriked-out&S,'#13#10'&Ooutlined&O,'#13#10'&Hshadow'

'&LH&Y2&YO cm&X2&C&"Times New Roman"&18This is big&RThis is &Bbold&B,'#13#10'&Iitalic&I,'#13#10'&Uunderlined&U,'#13#10'&Edouble underlined&E,&Sstriked-out&S,'#13#10'&Ooutlined&O,'#13#10'&Hshadow'
'&LH&Y2&YO cm&X2  &C&"Times New Roman"&18This is big&RThis is &Bbold&B,'#13#10'&Iitalic&I,'#13#10'&Uunderlined&U,'#13#10'&Edouble underlined&E,striked-out,'#13#10'&Ooutlined&O,'#13#10'&Hshadow'
}


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_PrintRanges(
  AFormat: TsSpreadsheetFormat; ANumSheets, ANumRanges: Integer; ASpaceInName: Boolean);
var
  tempFile: String;
  i, j: Integer;
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  rng: TsCellRange;
  sheetname: String;
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    for i:= 1 to ANumSheets do
    begin
      sheetname := PageLayoutSheet + IfThen(ASpaceInName, ' ', '') + IntToStr(i);
      MyWorksheet := MyWorkbook.AddWorksheet(sheetname);
      for j:=1 to ANumRanges do
        MyWorksheet.PageLayout.AddPrintRange(SollRanges[j]);
    end;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    CheckEquals(ANumSheets, MyWorkbook.GetWorksheetCount, 'Worksheet count mismatch');
    for i := 1 to ANumSheets do
    begin
      MyWorksheet := MyWorkbook.GetWorksheetByIndex(i-1);
      CheckEquals(ANumRanges, MyWorksheet.PageLayout.NumPrintRanges, 'Print range count mismatch');
      for j:=1 to ANumRanges do
      begin
        rng := MyWorksheet.PageLayout.GetPrintRange(j-1);
        CheckEquals(SollRanges[j].Row1, rng.Row1, Format('Row1 mismatch at i=%d, j=%d', [i, j]));
        CheckEquals(SollRanges[j].Row2, rng.Row2, Format('Row2 mismatch at i=%d, j=%d', [i, j]));
        CheckEquals(SollRanges[j].Col1, rng.Col1, Format('Col1 mismatch at i=%d, j=%d', [i, j]));
        CheckEquals(SollRanges[j].Col2, rng.Col2, Format('Col2 mismatch at i=%d, j=%d', [i, j]));
      end;
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_RepeatedColRows(
  AFormat: TsSpreadsheetFormat; AFirstCol, ALastCol, AFirstRow, ALastRow: Integer);
var
  tempFile: String;
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  sheetname: String;
  r, c: Cardinal;
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    sheetname := PageLayoutSheet;
    MyWorksheet := MyWorkbook.AddWorksheet(sheetname);
    for r := 0 to 10 do
      for c := 0 to 10 do
        MyWorksheet.WriteNumber(r, c, r*100+c);
    MyWorksheet.PageLayout.SetRepeatedRows(Cardinal(AFirstRow), Cardinal(ALastRow));
    MyWorksheet.PageLayout.SetRepeatedCols(Cardinal(AFirstCol), Cardinal(ALastCol));
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    MyWorksheet := MyWorkbook.GetWorksheetByName(sheetname);
    CheckEquals(Cardinal(AFirstRow), MyWorksheet.Pagelayout.RepeatedRows.FirstIndex, 'First repeated row index mismatch');
    CheckEquals(Cardinal(ALastRow), MyWorksheet.PageLayout.RepeatedRows.LastIndex, 'Last repeated row index mismatch');
    CheckEquals(Cardinal(AFirstCol), MyWorksheet.PageLayout.RepeatedCols.FirstIndex, 'First repeated col index mismatch');
    CheckEquals(Cardinal(ALastCol), MyWorksheet.PageLayout.RepeatedCols.LastIndex, 'Last repeated col index mismatch');
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


{ Tests for BIFF2 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel2, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterRegions_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel2, 1, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterRegions_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel2, 2, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterRegions_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel2, 3, 6);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel2, 1, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel2, 2, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel2, 3, 7);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterFontSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel2, 1, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterFontSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel2, 2, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF2_HeaderFooterFontSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel2, 3, 8);
end;


{ Tests for BIFF5 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel5, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 2);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_WidthToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_WidthToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_WidthToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeightToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeightToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeightToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 4);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageNumber_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageNumber_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PageNumber_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 5);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterRegions_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterRegions_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterRegions_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 6);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 7);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterFontSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterFontSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterFontSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 8);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterFontColor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel5, 1, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterFontColor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 2, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_HeaderFooterFontColor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel5, 3, 9);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_1sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel5, 1, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_1sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel5, 1, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_2sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel5, 2, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_2sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel5, 2, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_1sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfExcel5, 1, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_1sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfExcel5, 1, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_2sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfExcel5, 2, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_PrintRanges_2sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfExcel5, 2, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedRow_0;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, -1, -1, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedRows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, -1, -1, 0, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedRows_1_3;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, -1, -1, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedCol_0;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, 0, 0, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedCols_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, 0, 1, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedCols_1_3;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, 1, 3, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedCol_0_Row_0;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, 0, 0, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF5_RepeatedCols_0_1_Rows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcel5, 0, 1, 0, 1);
end;


{ Tests for BIFF8 file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcel8, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 2);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_WidthToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_WidthToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_WidthToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeightToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeightToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeightToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 4);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageNumber_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageNumber_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PageNumber_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 5);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterRegions_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterRegions_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterRegions_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 6);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 7);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterFontSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterFontSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterFontSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 8);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterFontColor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcel8, 1, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterFontColor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 2, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_HeaderFooterFontColor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcel8, 3, 9);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_1sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel8, 1, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_1sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel8, 1, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_2sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel8, 2, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_2sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcel8, 2, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_1sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfExcel8, 1, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_1sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfExcel8, 1, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_2sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfExcel8, 2, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_PrintRanges_2sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfExcel8, 2, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedRow_0;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, -1, -1, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedRows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, -1, -1, 0, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedRows_1_3;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, -1, -1, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedCol_0;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, 0, 0, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedCols_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, 0, 1, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedCols_1_3;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, 1, 3, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedCol_0_Row_0;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, 0, 0, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_BIFF8_RepeatedCols_0_1_Rows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcel8, 0, 1, 0, 1);
end;


{ Tests for OOXML file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfOOXML, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 2);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_WidthToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_WidthToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_WidthToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeightToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeightToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeightToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 4);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageNumber_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageNumber_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PageNumber_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 5);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterRegions_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterRegions_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterRegions_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 6);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 7);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterFontSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterFontSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterFontSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 8);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterFontColor_1sheet;
begin
  TestWriteRead_PageLayout(sfOOXML, 1, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterFontColor_2sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 2, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_HeaderFooterFontColor_3sheets;
begin
  TestWriteRead_PageLayout(sfOOXML, 3, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_1sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOOXML, 1, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_1sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOOXML, 1, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_2sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOOXML, 2, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_2sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOOXML, 2, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_1sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfOOXML, 1, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_1sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfOOXML, 1, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_2sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfOOXML, 2, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_PrintRanges_2sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfOOXML, 2, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedRow_0;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, -1, -1, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedRows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, -1, -1, 0, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedRows_1_3;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, -1, -1, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedCol_0;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, 0, 0, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedCols_0_1;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, 0, 1, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedCols_1_3;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, 1, 3, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedCol_0_Row_0;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, 0, 0, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_OOXML_RepeatedCols_0_1_Rows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfOOXML, 0, 1, 0, 1);
end;


{ Tests for Excdl2003/XML file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfExcelXML, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfExcelXML, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfExcelXML, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfExcelXML, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_XML_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfExcelXML, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfExcelXML, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfExcelXML, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfExcelXML, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfExcelXML, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfExcelXML, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfExcelXML, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfExcelXML, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 2);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_WidthToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_WidthToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_WidthToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeightToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeightToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeightToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 4);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageNumber_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageNumber_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PageNumber_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 5);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterRegions_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterRegions_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterRegions_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 6);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 7);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterFontSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterFontSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterFontSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 8);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterFontColor_1sheet;
begin
  TestWriteRead_PageLayout(sfExcelXML, 1, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterFontColor_2sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 2, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_HeaderFooterFontColor_3sheets;
begin
  TestWriteRead_PageLayout(sfExcelXML, 3, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_1sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 1, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_1sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 1, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_2sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 2, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_2sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 2, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_1sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 1, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_1sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 1, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_2sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 2, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_PrintRanges_2sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfExcelXML, 2, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedRow_0;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, -1, -1, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedRows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, -1, -1, 0, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedRows_1_3;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, -1, -1, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedCol_0;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, 0, 0, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedCols_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, 0, 1, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedCols_1_3;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, 1, 3, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedCol_0_Row_0;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, 0, 0, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_XML_RepeatedCols_0_1_Rows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfExcelXML, 0, 1, 0, 1);
end;


{ Tests for Open Document file format }

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_1sheet_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 1, 3);
end;


procedure TSpreadWriteReadPagelayoutTests.TestWriteRead_ODS_PageMargins_2sheets_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_2sheets_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_2sheets_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_2sheets_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 2, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_0;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_1;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_2;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageMargins_3sheets_3;
begin
  TestWriteRead_PageMargins(sfOpenDocument, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageOrientation_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageOrientation_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageOrientation_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 0);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PaperSize_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PaperSize_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PaperSize_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 1);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_ScalingFactor_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_ScalingFactor_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 2);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_ScalingFactor_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 2);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_WidthToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_WidthToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_WidthToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 3);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeightToPages_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeightToPages_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 4);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeightToPages_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 4);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageNumber_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageNumber_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 5);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PageNumber_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 5);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterRegions_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterRegions_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 6);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterRegions_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 6);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 7);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 7);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterFontSymbols_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterFontSymbols_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 8);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterFontSymbols_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 8);
end;


procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterFontColor_1sheet;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 1, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterFontColor_2sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 2, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_HeaderFooterFontColor_3sheets;
begin
  TestWriteRead_PageLayout(sfOpenDocument, 3, 9);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_1sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 1, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_1sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 1, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_2sheet_1Range_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 2, 1, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_2sheet_2Ranges_NoSpace;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 2, 2, false);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_1sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 1, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_1sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 1, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_2sheet_1Range_Space;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 2, 1, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_PrintRanges_2sheet_2Ranges_Space;
begin
  TestWriteRead_PrintRanges(sfOpenDocument, 2, 2, true);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedRow_0;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, -1, -1, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedRows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, -1, -1, 0, 1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedRows_1_3;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, -1, -1, 1, 3);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedCol_0;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, 0, 0, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedCols_0_1;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, 0, 1, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedCols_1_3;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, 1, 3, -1, -1);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedCol_0_Row_0;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, 0, 0, 0, 0);
end;

procedure TSpreadWriteReadPageLayoutTests.TestWriteRead_ODS_RepeatedCols_0_1_Rows_0_1;
begin
  TestWriteRead_RepeatedColRows(sfOpenDocument, 0, 1, 0, 1);
end;



initialization
  RegisterTest(TSpreadWriteReadPageLayoutTests);

end.

