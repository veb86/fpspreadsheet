{ Tests for insertion and deletion of columns and rows
  This unit test is writing out to and reading back from files.
}

unit insertdeletetests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff8, {and a project requirement for lclbase for utf8 handling}
  testsutility;

type
  TInsDelTestDataItem = record
    Layout: string;
    InsertCol: Integer;
    InsertRow: Integer;
    DeleteCol: Integer;
    DeleteRow: Integer;
    Formula: String;
    SollFormula: String;
    SharedFormulaRowCount: Integer;    // Size of shared formula block before insert/delete
    SharedFormulaColCount: Integer;
    SharedFormulaBaseCol_After: Integer;   // Position of shared formula base after insert/delete
    SharedFormulaBaseRow_After: Integer;
    SharedFormulaRowCount_After: Integer;  // Size of shared formula block after insert/delete
    SharedFormulaColCount_After: Integer;
    MergedColCount: Integer;      // size of merged block before insert/delete
    MergedRowCount: Integer;
    MergedColCount_After: Integer;  // size of merged block after insert/delete
    MergedRowCount_After: Integer;
    SollLayout: String;
  end;

var
  InsDelTestData: array[0..52] of TInsDelTestDataItem;

  procedure InitTestData;

type
  { TSpreadWriteReadInsertColRowTests }
  TSpreadWriteRead_InsDelColRow_Tests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_InsDelColRow(ATestIndex: Integer;
      AFormat: TsSpreadsheetFormat);

  published

    // *** Excel 8 tests ***

    // Writes out simple cell layout and inserts columns
    procedure TestWriteRead_InsDelColRow_0_BIFF8;     // before first
    procedure TestWriteRead_InsDelColRow_1_BIFF8;     // middle
    procedure TestWriteRead_InsDelColRow_2_BIFF8;     // before last
    // Writes out simple cell layout and deletes columns
    procedure TestWriteRead_InsDelColRow_3_BIFF8;     // first
    procedure TestWriteRead_InsDelColRow_4_BIFF8;     // middle
    procedure TestWriteRead_InsDelColRow_5_BIFF8;     // last
    // Writes out simple cell layout and inserts rows
    procedure TestWriteRead_InsDelColRow_6_BIFF8;     // before first
    procedure TestWriteRead_InsDelColRow_7_BIFF8;     // middle
    procedure TestWriteRead_InsDelColRow_8_BIFF8;     // before last
    // Writes out simple cell layout and deletes rows
    procedure TestWriteRead_InsDelColRow_9_BIFF8;     // first
    procedure TestWriteRead_InsDelColRow_10_BIFF8;    // middle
    procedure TestWriteRead_InsDelColRow_11_BIFF8;    // last

    // Writes out cell layout with formula and inserts columns
    procedure TestWriteRead_InsDelColRow_12_BIFF8;    // before formula cell
    procedure TestWriteRead_InsDelColRow_13_BIFF8;    // after formula cell
    // Writes out cell layout with formula and inserts rows
    procedure TestWriteRead_InsDelColRow_14_BIFF8;    // before formula cell
    procedure TestWriteRead_InsDelColRow_15_BIFF8;    // after formula cell
    // Writes out cell layout with formula and deletes columns
    procedure TestWriteRead_InsDelColRow_16_BIFF8;    // before formula cell
    procedure TestWriteRead_InsDelColRow_17_BIFF8;    // after formula cell
    procedure TestWriteRead_InsDelColRow_18_BIFF8;    // cell in formula
    // Writes out cell layout with formula and deletes rows
    procedure TestWriteRead_InsDelColRow_19_BIFF8;    // before formula cell
    procedure TestWriteRead_InsDelColRow_20_BIFF8;    // after formula cell
    procedure TestWriteRead_InsDelColRow_21_BIFF8;    // cell in formula

    // Writes out cell layout with shared formula
    procedure TestWriteRead_InsDelColRow_22_BIFF8;    // no insert/delete; just test shared formula
    // ... and inserts columns
    procedure TestWriteRead_InsDelColRow_23_BIFF8;    // column before shared formula cells
    procedure TestWriteRead_InsDelColRow_24_BIFF8;    // column after shared formula cells
    procedure TestWriteRead_InsDelColRow_25_BIFF8;    // column through cells addressed by shared formula
    procedure TestWriteRead_InsDelColRow_26_BIFF8;    // column through shared formula block
    procedure TestWriteRead_InsDelColRow_27_BIFF8;    // column behind shared formula block
    // ... and inserts rows
    procedure TestWriteRead_InsDelColRow_28_BIFF8;    // row before shared formula
    procedure TestWriteRead_InsDelColRow_29_BIFF8;    // row after shared formula cells
    procedure TestWriteRead_InsDelColRow_30_BIFF8;    // row through shared formula block
    procedure TestWriteRead_InsDelColRow_31_BIFF8;    // row below shared formula block
    // ... and deletes columns
    procedure TestWriteRead_InsDelColRow_32_BIFF8;    // column before shared formula cells
    procedure TestWriteRead_InsDelColRow_33_BIFF8;    // column after shared formula cells
    procedure TestWriteRead_InsDelColRow_34_BIFF8;    // column in shared formula block, no formula cell
    procedure TestWriteRead_InsDelColRow_35_BIFF8;    // column used in shared formula
    // ... and deletes rows
    procedure TestWriteRead_InsDelColRow_36_BIFF8;    // row above shared formula cells
    procedure TestWriteRead_InsDelColRow_37_BIFF8;    // row below shared formula cells
    procedure TestWriteRead_InsDelColRow_38_BIFF8;    // row through shared formula block
    procedure TestWriteRead_InsDelColRow_39_BIFF8;    // row with cell used in shared formula

    // Writes out cell layout with merged cells
    procedure TestWriteRead_InsDelColRow_40_BIFF8;    // no insert/delete; just test merged block
    // ... and inserts columns
    procedure TestWriteRead_InsDelColRow_41_BIFF8;    // column before merged block
    procedure TestWriteRead_InsDelColRow_42_BIFF8;    // column through merged block
    procedure TestWriteRead_InsDelColRow_43_BIFF8;    // column after merged block
    // ... and inserts rows
    procedure TestWriteRead_InsDelColRow_44_BIFF8;    // row before merged block
    procedure TestWriteRead_InsDelColRow_45_BIFF8;    // row through merged block
    procedure TestWriteRead_InsDelColRow_46_BIFF8;    // row after merged block
    // ... and deletes columns
    procedure TestWriteRead_InsDelColRow_47_BIFF8;    // column before merged block
    procedure TestWriteRead_InsDelColRow_48_BIFF8;    // column through merged block
    procedure TestWriteRead_InsDelColRow_49_BIFF8;    // column after merged block
    // ... and deletes rows
    procedure TestWriteRead_InsDelColRow_50_BIFF8;    // row before merged block
    procedure TestWriteRead_InsDelColRow_51_BIFF8;    // row through merged block
    procedure TestWriteRead_InsDelColRow_52_BIFF8;    // row after merged block

    // *** OOXML tests ***

    // Writes out simple cell layout and inserts columns
    procedure TestWriteRead_InsDelColRow_0_OOXML;     // before first
    procedure TestWriteRead_InsDelColRow_1_OOXML;     // middle
    procedure TestWriteRead_InsDelColRow_2_OOXML;     // before last
    // Writes out simple cell layout and deletes columns
    procedure TestWriteRead_InsDelColRow_3_OOXML;     // first
    procedure TestWriteRead_InsDelColRow_4_OOXML;     // middle
    procedure TestWriteRead_InsDelColRow_5_OOXML;     // last
    // Writes out simple cell layout and inserts rows
    procedure TestWriteRead_InsDelColRow_6_OOXML;     // before first
    procedure TestWriteRead_InsDelColRow_7_OOXML;     // middle
    procedure TestWriteRead_InsDelColRow_8_OOXML;     // before last
    // Writes out simple cell layout and deletes rows
    procedure TestWriteRead_InsDelColRow_9_OOXML;     // first
    procedure TestWriteRead_InsDelColRow_10_OOXML;    // middle
    procedure TestWriteRead_InsDelColRow_11_OOXML;    // last

    // Writes out cell layout with formula and inserts columns
    procedure TestWriteRead_InsDelColRow_12_OOXML;    // before formula cell
    procedure TestWriteRead_InsDelColRow_13_OOXML;    // after formula cell
    // Writes out cell layout with formula and inserts rows
    procedure TestWriteRead_InsDelColRow_14_OOXML;    // before formula cell
    procedure TestWriteRead_InsDelColRow_15_OOXML;    // after formula cell
    // Writes out cell layout with formula and deletes columns
    procedure TestWriteRead_InsDelColRow_16_OOXML;    // before formula cell
    procedure TestWriteRead_InsDelColRow_17_OOXML;    // after formula cell
    procedure TestWriteRead_InsDelColRow_18_OOXML;    // cell in formula
    // Writes out cell layout with formula and deletes rows
    procedure TestWriteRead_InsDelColRow_19_OOXML;    // before formula cell
    procedure TestWriteRead_InsDelColRow_20_OOXML;    // after formula cell
    procedure TestWriteRead_InsDelColRow_21_OOXML;    // cell in formula

    // Writes out cell layout with shared formula
    procedure TestWriteRead_InsDelColRow_22_OOXML;    // no insert/delete; just test shared formula
    // ... and inserts columns
    procedure TestWriteRead_InsDelColRow_23_OOXML;    // column before shared formula cells
    procedure TestWriteRead_InsDelColRow_24_OOXML;    // column after shared formula cells
    procedure TestWriteRead_InsDelColRow_25_OOXML;    // column through cells addressed by shared formula
    procedure TestWriteRead_InsDelColRow_26_OOXML;    // column through shared formula block
    procedure TestWriteRead_InsDelColRow_27_OOXML;    // column behind shared formula block
    // ... and inserts rows
    procedure TestWriteRead_InsDelColRow_28_OOXML;    // row before shared formula
    procedure TestWriteRead_InsDelColRow_29_OOXML;    // row after shared formula cells
    procedure TestWriteRead_InsDelColRow_30_OOXML;    // row through shared formula block
    procedure TestWriteRead_InsDelColRow_31_OOXML;    // row below shared formula block
    // ... and deletes columns
    procedure TestWriteRead_InsDelColRow_32_OOXML;    // column before shared formula cells
    procedure TestWriteRead_InsDelColRow_33_OOXML;    // column after shared formula cells
    procedure TestWriteRead_InsDelColRow_34_OOXML;    // column in shared formula block, no formula cell
    procedure TestWriteRead_InsDelColRow_35_OOXML;    // column used in shared formula
    // ... and deletes rows
    procedure TestWriteRead_InsDelColRow_36_OOXML;    // row above shared formula cells
    procedure TestWriteRead_InsDelColRow_37_OOXML;    // row below shared formula cells
    procedure TestWriteRead_InsDelColRow_38_OOXML;    // row through shared formula block
    procedure TestWriteRead_InsDelColRow_39_OOXML;    // row with cell used in shared formula

    // Writes out cell layout with merged cells
    procedure TestWriteRead_InsDelColRow_40_OOXML;    // no insert/delete; just test merged block
    // ... and inserts columns
    procedure TestWriteRead_InsDelColRow_41_OOXML;    // column before merged block
    procedure TestWriteRead_InsDelColRow_42_OOXML;    // column through merged block
    procedure TestWriteRead_InsDelColRow_43_OOXML;    // column after merged block
    // ... and inserts rows
    procedure TestWriteRead_InsDelColRow_44_OOXML;    // row before merged block
    procedure TestWriteRead_InsDelColRow_45_OOXML;    // row through merged block
    procedure TestWriteRead_InsDelColRow_46_OOXML;    // row after merged block
    // ... and deletes columns
    procedure TestWriteRead_InsDelColRow_47_OOXML;    // column before merged block
    procedure TestWriteRead_InsDelColRow_48_OOXML;    // column through merged block
    procedure TestWriteRead_InsDelColRow_49_OOXML;    // column after merged block
    // ... and deletes rows
    procedure TestWriteRead_InsDelColRow_50_OOXML;    // row before merged block
    procedure TestWriteRead_InsDelColRow_51_OOXML;    // row through merged block
    procedure TestWriteRead_InsDelColRow_52_OOXML;    // row after merged block

    // *** OpenDocument tests ***

    // Writes out simple cell layout and inserts columns
    procedure TestWriteRead_InsDelColRow_0_ODS;     // before first
    procedure TestWriteRead_InsDelColRow_1_ODS;     // middle
    procedure TestWriteRead_InsDelColRow_2_ODS;     // before last
    // Writes out simple cell layout and deletes columns
    procedure TestWriteRead_InsDelColRow_3_ODS;     // first
    procedure TestWriteRead_InsDelColRow_4_ODS;     // middle
    procedure TestWriteRead_InsDelColRow_5_ODS;     // last
    // Writes out simple cell layout and inserts rows
    procedure TestWriteRead_InsDelColRow_6_ODS;     // before first
    procedure TestWriteRead_InsDelColRow_7_ODS;     // middle
    procedure TestWriteRead_InsDelColRow_8_ODS;     // before last
    // Writes out simple cell layout and deletes rows
    procedure TestWriteRead_InsDelColRow_9_ODS;     // first
    procedure TestWriteRead_InsDelColRow_10_ODS;    // middle
    procedure TestWriteRead_InsDelColRow_11_ODS;    // last

    // Writes out cell layout with formula and inserts columns
    procedure TestWriteRead_InsDelColRow_12_ODS;    // before formula cell
    procedure TestWriteRead_InsDelColRow_13_ODS;    // after formula cell
    // Writes out cell layout with formula and inserts rows
    procedure TestWriteRead_InsDelColRow_14_ODS;    // before formula cell
    procedure TestWriteRead_InsDelColRow_15_ODS;    // after formula cell
    // Writes out cell layout with formula and deletes columns
    procedure TestWriteRead_InsDelColRow_16_ODS;    // before formula cell
    procedure TestWriteRead_InsDelColRow_17_ODS;    // after formula cell
    procedure TestWriteRead_InsDelColRow_18_ODS;    // cell in formula
    // Writes out cell layout with formula and deletes rows
    procedure TestWriteRead_InsDelColRow_19_ODS;    // before formula cell
    procedure TestWriteRead_InsDelColRow_20_ODS;    // after formula cell
    procedure TestWriteRead_InsDelColRow_21_ODS;    // cell in formula

    // Writes out cell layout with shared formula
    procedure TestWriteRead_InsDelColRow_22_ODS;    // no insert/delete; just test shared formula
    // ... and inserts columns
    procedure TestWriteRead_InsDelColRow_23_ODS;    // column before shared formula cells
    procedure TestWriteRead_InsDelColRow_24_ODS;    // column after shared formula cells
    procedure TestWriteRead_InsDelColRow_25_ODS;    // column through cells addressed by shared formula
    procedure TestWriteRead_InsDelColRow_26_ODS;    // column through shared formula block
    procedure TestWriteRead_InsDelColRow_27_ODS;    // column behind shared formula block
    // ... and inserts rows
    procedure TestWriteRead_InsDelColRow_28_ODS;    // row before shared formula
    procedure TestWriteRead_InsDelColRow_29_ODS;    // row after shared formula cells
    procedure TestWriteRead_InsDelColRow_30_ODS;    // row through shared formula block
    procedure TestWriteRead_InsDelColRow_31_ODS;    // row below shared formula block
    // ... and deletes columns
    procedure TestWriteRead_InsDelColRow_32_ODS;    // column before shared formula cells
    procedure TestWriteRead_InsDelColRow_33_ODS;    // column after shared formula cells
    procedure TestWriteRead_InsDelColRow_34_ODS;    // column in shared formula block, no formula cell
    procedure TestWriteRead_InsDelColRow_35_ODS;    // column used in shared formula
    // ... and deletes rows
    procedure TestWriteRead_InsDelColRow_36_ODS;    // row above shared formula cells
    procedure TestWriteRead_InsDelColRow_37_ODS;    // row below shared formula cells
    procedure TestWriteRead_InsDelColRow_38_ODS;    // row through shared formula block
    procedure TestWriteRead_InsDelColRow_39_ODS;    // row with cell used in shared formula

    // Writes out cell layout with merged cells
    procedure TestWriteRead_InsDelColRow_40_ODS;    // no insert/delete; just test merged block
    // ... and inserts columns
    procedure TestWriteRead_InsDelColRow_41_ODS;    // column before merged block
    procedure TestWriteRead_InsDelColRow_42_ODS;    // column through merged block
    procedure TestWriteRead_InsDelColRow_43_ODS;    // column after merged block
    // ... and inserts rows
    procedure TestWriteRead_InsDelColRow_44_ODS;    // row before merged block
    procedure TestWriteRead_InsDelColRow_45_ODS;    // row through merged block
    procedure TestWriteRead_InsDelColRow_46_ODS;    // row after merged block
    // ... and deletes columns
    procedure TestWriteRead_InsDelColRow_47_ODS;    // column before merged block
    procedure TestWriteRead_InsDelColRow_48_ODS;    // column through merged block
    procedure TestWriteRead_InsDelColRow_49_ODS;    // column after merged block
    // ... and deletes rows
    procedure TestWriteRead_InsDelColRow_50_ODS;    // row before merged block
    procedure TestWriteRead_InsDelColRow_51_ODS;    // row through merged block
    procedure TestWriteRead_InsDelColRow_52_ODS;    // row after merged block
  end;

implementation

uses
  StrUtils;

const
  InsertColRowSheet = 'InsertDelete_ColumnsRows';

procedure InitTestData;
var
  i: Integer;
begin
  for i := 0 to High(InsDelTestData) do
    with InsDelTestData[i] do
    begin
      Layout := '';
      InsertCol := -1;
      InsertRow := -1;
      DeleteCol := -1;
      DeleteRow := -1;
      Formula := '';
      SollFormula := '';
      SharedFormulaColCount := 0;
      SharedFormulaRowCount := 0;
      SharedFormulaBaseCol_After := -1;
      SharedFormulaBaseRow_After := -1;
      SharedFormulaColCount_After := 0;
      SharedFormulaRowCount_After := 0;
      MergedColCount := 0;
      MergedRowCount := 0;
    end;

  { ---------------------------------------------------------------------------}
  {  Simple layouts                                                            }
  { ---------------------------------------------------------------------------}

  // Insert a column before col 0
  with InsDelTestData[0] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    InsertCol := 0;
    SollLayout := ' 12345678|'+
                  ' 23456789|'+
                  ' 34567890|'+
                  ' 45678901';
  end;

  // Insert a column before col 2
  with InsDelTestData[1] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    InsertCol := 2;
    SollLayout := '12 345678|'+
                  '23 456789|'+
                  '34 567890|'+
                  '45 678901';
  end;

  // Insert a column before last col
  with InsDelTestData[2] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    InsertCol := 7;
    SollLayout := '1234567 8|'+
                  '2345678 9|'+
                  '3456789 0|'+
                  '4567890 1';
  end;

  // Delete column 0
  with InsDelTestData[3] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    DeleteCol := 0;
    SollLayout := '2345678|'+
                  '3456789|'+
                  '4567890|'+
                  '5678901';
  end;

  // Delete column 2
  with InsDelTestData[4] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    DeleteCol := 2;
    SollLayout := '1245678|'+
                  '2356789|'+
                  '3467890|'+
                  '4578901';
  end;

  // Delete last column
  with InsDelTestData[5] do begin
    Layout := '12345678|'+
              '23456789|'+
              '34567890|'+
              '45678901';
    DeleteCol := 7;
    SollLayout := '1234567|'+
                  '2345678|'+
                  '3456789|'+
                  '4567890';
  end;

  // Insert a ROW before row 0
  with InsDelTestData[6] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    InsertRow := 0;
    SollLayout := '     |'+
                  '12345|'+
                  '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Insert a ROW before row 2
  with InsDelTestData[7] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    InsertRow := 2;
    SollLayout := '12345|'+
                  '23456|'+
                  '     |'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Insert a ROW before last row
  with InsDelTestData[8] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    InsertRow := 5;
    SollLayout := '12345|'+
                  '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '     |'+
                  '67890|';
  end;

  // Delete the first row
  with InsDelTestData[9] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    DeleteRow := 0;
    SollLayout := '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Delete row #2
  with InsDelTestData[10] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    DeleteRow := 2;
    SollLayout := '12345|'+
                  '23456|'+
                  '45678|'+
                  '56789|'+
                  '67890|';
  end;

  // Delete last row
  with InsDelTestData[11] do begin
    Layout := '12345|'+
              '23456|'+
              '34567|'+
              '45678|'+
              '56789|'+
              '67890|';
    DeleteRow := 5;
    SollLayout := '12345|'+
                  '23456|'+
                  '34567|'+
                  '45678|'+
                  '56789';
  end;

  { ---------------------------------------------------------------------------}
  {  Layouts with formula                                                      }
  { ---------------------------------------------------------------------------}

  // Insert a column before #1, i.e. before formula cell
  with InsDelTestData[12] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertCol := 1;
    Formula := 'C3';
    SollFormula := 'D3';           // col index increases due to inserted col
    SollLayout := '1 2345678|'+
                  '2 3456789|'+
                  '3 4565890|'+
                  '4 5678901|'+
                  '5 6789012|'+
                  '6 7890123';
  end;

  // Insert a column before #3, i.e. after formula cell
  with InsDelTestData[13] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertCol := 3;
    Formula := 'C3';
    SollFormula := 'C3';           // no change of cell because insertion is behind
    SollLayout := '123 45678|'+
                  '234 56789|'+
                  '345 65890|'+
                  '456 78901|'+
                  '567 89012|'+
                  '678 90123';
  end;

  // Insert a row before #1, i.e. before formula cell
  with InsDelTestData[14] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertRow := 1;
    Formula := 'E4';
    SollFormula := 'E5';         // row index increaes due to inserted row
    SollLayout := '12345678|'+
                  '        |'+
                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
                  '56789012|'+
                  '67890123';
  end;

  // Insert a row before #4, i.e. after formula cell
  with InsDelTestData[15] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    InsertRow := 5;
    Formula := 'E4';
    SollFormula := 'E4';         // row index not changed dur to insert after cell
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
                  '56789012|'+
                  '        |'+
                  '67890123';
  end;

  // Deletes column #1, i.e. before formula cell
  with InsDelTestData[16] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteCol := 1;
    Formula := 'C3';
    SollFormula := 'B3';           // col index decreases due to delete before cell
    SollLayout := '1345678|'+
                  '2456789|'+
                  '3565890|'+
                  '4678901|'+
                  '5789012|'+
                  '6890123';
  end;

  // Deletes column #5, i.e. after formula cell
  with InsDelTestData[17] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteCol := 5;
    Formula := 'C3';
    SollFormula := 'C3';         // col index unchanged due to deleted after cell
    SollLayout := '1234578|'+
                  '2345689|'+
                  '3456590|'+
                  '4567801|'+
                  '5678912|'+
                  '6789023';
  end;

  // Deletes column #2, i.e. cell appearing in formula is gone --> #REF! error
  with InsDelTestData[18] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteCol := 2;
    Formula := 'C3';
    SollFormula := '#REF!';         // col index unchanged due to deletion after cell
    SollLayout := '1245678|'+
                  '2356789|'+
                  '346E890|'+    // "E" = error
                  '4578901|'+
                  '5689012|'+
                  '6790123';
  end;

  // Deletes row #1, i.e. before formula cell
  with InsDelTestData[19] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteRow := 1;
    Formula := 'E4';
    SollFormula := 'E3';           // row index decreases due to delete before cell
    SollLayout := '12345678|'+
//                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
                  '56789012|'+
                  '67890123';
  end;

  // Deletes row #4, i.e. after formula cell
  with InsDelTestData[20] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteRow := 4;
    Formula := 'E4';
    SollFormula := 'E4';           // row index unchanged (delete is after cell)
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34568890|'+
                  '45678901|'+
//                  '56789012|'+
                  '67890123';
  end;

  // Deletes row #2, i.e. row containing cell used in formula --> #REF! error!
  with InsDelTestData[21] do begin
    Layout := '12345678|'+
              '23456789|'+
              '3456F890|'+                   // "F" = Formula in row 2, col 4
              '45678901|'+
              '56789012|'+
              '67890123';
    DeleteRow := 3;
    Formula := 'E4';
    SollFormula := '#REF!';
    SollLayout := '12345678|'+
                  '23456789|'+
                  '3456E890|'+    // "E" = error
//                  '45678901|'+
                  '56789012|'+
                  '67890123';
  end;

  { ---------------------------------------------------------------------------}
  {  Layouts with shared formula                                                      }
  { ---------------------------------------------------------------------------}

  // No insert/delete, just to test the shared formula
  with InsDelTestData[22] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A3-$B$2,B3-$B$2;'+
                   'A4-$B$2,B4-$B$2;'+
                   'A5-$B$2,B5-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34501890|'+
                  '45612901|'+
                  '56723012|'+
                  '67890123';
  end;

  // Insert column before any cell referred to by the shared formula (col = 0)
  with InsDelTestData[23] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 0;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 4;   // Position of shared formula after insert op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'B3-$C$2,C3-$C$2;'+   // all column indexes increase by 1 due to added col in front
                   'B4-$C$2,C4-$C$2;'+
                   'B5-$C$2,C5-$C$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := ' 12345678|'+
                  ' 23456789|'+
                  ' 34501890|'+
                  ' 45612901|'+
                  ' 56723012|'+
                  ' 67890123';
  end;

  // Insert column after last cell addressed by the shared formula
  with InsDelTestData[24] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 7;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A3-$B$2,B3-$B$2;'+    // formulas unchanged by insert
                   'A4-$B$2,B4-$B$2;'+
                   'A5-$B$2,B5-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '1234567 8|'+
                  '2345678 9|'+
                  '3450189 0|'+
                  '4561290 1|'+
                  '5672301 2|'+
                  '6789012 3';
  end;

  // Insert column between cells referred to by the shared formula (col = 1)
  with InsDelTestData[25] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 1;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 4;
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A3-$C$2,C3-$C$2;'+
                   'A4-$C$2,C4-$C$2;'+
                   'A5-$C$2,C5-$C$2';
    SollLayout := '1 2345678|'+
                  '2 3456789|'+
                  '3 4501890|'+
                  '4 5612901|'+
                  '5 6723012|'+
                  '6 7890123';
  end;

  // Insert column to run through formula block (col = 4)
  with InsDelTestData[26] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 4;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 3;
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A3-$B$2,'',B3-$B$2;'+
                   'A4-$B$2,'',B4-$B$2;'+
                   'A5-$B$2,'',B5-$B$2';
    SollLayout := '1234 5678|'+
                  '2345 6789|'+
                  '3450 1890|'+
                  '4561 2901|'+
                  '5672 3012|'+
                  '6789 0123';
  end;

  // Insert column behind shared formula block (col = 7)
  with InsDelTestData[27] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 7;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A3-$B$2,B3-$B$2;'+
                   'A4-$B$2,B4-$B$2;'+
                   'A5-$B$2,B5-$B$2';
    SollLayout := '1234567 8|'+
                  '2345678 9|'+
                  '3450189 0|'+
                  '4561290 1|'+
                  '5672301 2|'+
                  '6789012 3';
  end;

  // Insert row before any cell referred to by the shared formula (row = 0)
  with InsDelTestData[28] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertRow := 0;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after insert op
    SharedFormulaBaseRow_After := 3;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A4-$B$3,B4-$B$3;'+   // all row indexes increase by 1 due to added row in front
                   'A5-$B$3,B5-$B$3;'+
                   'A6-$B$3,B6-$B$3';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '        |'+
                  '12345678|'+
                  '23456789|'+
                  '34501890|'+
                  '45612901|'+
                  '56723012|'+
                  '67890123';
  end;

  // Insert row through cells referred to by the shared formula (row = 2)
  with InsDelTestData[29] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertRow := 2;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after insert op
    SharedFormulaBaseRow_After := 3;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A4-$B$2,B4-$B$2;'+   // row 3 --> 4
                   'A5-$B$2,B5-$B$2;'+
                   'A6-$B$2,B6-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+
                  '23456789|'+
                  '        |'+
                  '34501890|'+
                  '45612901|'+
                  '56723012|'+
                  '67890123';
  end;

  // Insert row through shared formula block (row = 3)
  with InsDelTestData[30] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertRow := 3;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after insert op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 4;
    SollFormula := 'A3-$B$2,B3-$B$2;'+
                   '       ,       ;'+
                   'A5-$B$2,B5-$B$2;'+
                   'A6-$B$2,B6-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34501890|'+
                  '        |'+
                  '45612901|'+
                  '56723012|'+
                  '67890123';
  end;

  // Insert row below shared formula block (row = 5)
  with InsDelTestData[31] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertRow := 5;
    Formula := 'A3-$B$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after insert op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A3-$B$2,B3-$B$2;'+
                   'A4-$B$2,B4-$B$2;'+
                   'A5-$B$2,B5-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34501890|'+
                  '45612901|'+
                  '56723012|'+
                  '        |'+
                  '67890123';
  end;

  // Delete column 0 (this is before any cell of the shared formula)
  with InsDelTestData[32] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteCol := 0;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 2;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'A3-$B$2,B3-$B$2;'+   // all column indexes decrease by 1 due to deleted col in front
                   'A4-$B$2,B4-$B$2;'+
                   'A5-$B$2,B5-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '2345678|'+
                  '3456789|'+
                  '4501890|'+
                  '5612901|'+
                  '6723012|'+
                  '7890123';
  end;

  // Delete last column (this is behind any cell of the shared formula)
  with InsDelTestData[33] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteCol := 7;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'B3-$C$2,C3-$C$2;'+   // all column indexes unchanged due to delete behind
                   'B4-$C$2,C4-$C$2;'+
                   'B5-$C$2,C5-$C$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '1234567|'+
                  '2345678|'+
                  '3450189|'+
                  '4561290|'+
                  '5672301|'+
                  '6789012';
  end;

  // Delete in shared formula block, but no formula cell affected
  with InsDelTestData[34] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteCol := 4;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 1;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'B3-$C$2;'+
                   'B4-$C$2;'+
                   'B5-$C$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '1234678|'+
                  '2345789|'+
                  '3450890|'+
                  '4561901|'+
                  '5672012|'+
                  '6789123';
  end;

  // Delete column used in shared formula base
  with InsDelTestData[35] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteCol := 1;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 2;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := '#REF!-$B$2,#REF!-$B$2;'+
                   '#REF!-$B$2,#REF!-$B$2;'+
                   '#REF!-$B$2,#REF!-$B$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '1345678|'+
                  '2456789|'+
                  '35EE890|'+
                  '46EE901|'+
                  '57EE012|'+
                  '6890123';
  end;

  // Delete row 0 (this is before any cell of the shared formula)
  with InsDelTestData[36] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteRow := 0;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 1;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'B2-$C$1,C2-$C$1;'+   // all row indexes decrease by 1 due to deleted row in front
                   'B3-$C$1,C3-$C$1;'+
                   'B4-$C$1,C4-$C$1';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '23456789|'+
                  '34501890|'+
                  '45612901|'+
                  '56723012|'+
                  '67890123';
  end;

  // Delete lsst row (this is below any cell of the shared formula)
  with InsDelTestData[37] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteRow := 5;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'B3-$C$2,C3-$C$2;'+   // all row indexes unchanged
                   'B4-$C$2,C4-$C$2;'+
                   'B5-$C$2,C5-$C$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34501890|'+
                  '45612901|'+
                  '56723012|';
  end;

  // Delete row through shared formula block
  with InsDelTestData[38] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteRow := 3;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 2;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 2;
    SollFormula := 'B3-$C$2,C3-$C$2;'+
                   'B4-$C$2,C4-$C$2;'+
                   'B5-$C$2,C5-$C$2';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+
                  '23456789|'+
                  '34501890|'+
               //   '45612901|'+
                  '56723012|'+
                  '67890123';;
  end;

  // Delete row through shared formula block
  with InsDelTestData[39] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345S 890|'+                   // "S" = shared formula (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteRow := 1;
    Formula := 'B3-$C$2';
    SharedFormulaColCount := 2;
    SharedFormulaRowCount := 3;
    SharedFormulaBaseCol_After := 3;   // Position of shared formula after delete op
    SharedFormulaBaseRow_After := 1;
    SharedFormulaColCount_After := 2;  // Size of shared formula block after insert op
    SharedFormulaRowCount_After := 3;
    SollFormula := 'B2-#REF!,C2-#REF!;'+       // Confirmed by Excel
                   'B3-#REF!,C3-#REF!;'+
                   'B4-#REF!,C4-#REF!';
      // comma-separated --> cells along row; semicolon separates rows
    SollLayout := '12345678|'+                 // Confirmed by Excel
               //   '23456789|'+
                  '345EE890|'+
                  '456EE901|'+
                  '567EE012|'+
                  '67890123';;
  end;

  { ---------------------------------------------------------------------------}
  {  Layouts with merged cells                                                 }
  { ---------------------------------------------------------------------------}

  // No insert/delete, just to test the merged block
  with InsDelTestData[40] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '12345678|'+
                  '23456789|'+
                  '345M 890|'+
                  '456  901|'+
                  '567  012|'+
                  '67890123';
  end;

  // Insert column before merged block
  with InsDelTestData[41] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 1;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '1 2345678|'+
                  '2 3456789|'+
                  '3 45M 890|'+
                  '4 56  901|'+
                  '5 67  012|'+
                  '6 7890123';
  end;

  // Insert column through merged block
  with InsDelTestData[42] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 4;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 3;
    SollLayout := '1234 5678|'+
                  '2345 6789|'+
                  '345M  890|'+
                  '456   901|'+
                  '567   012|'+
                  '6789 0123';
  end;

  // Insert column behind merged block
  with InsDelTestData[43] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertCol := 7;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '1234567 8|'+
                  '2345678 9|'+
                  '345M 89 0|'+
                  '456  90 1|'+
                  '567  01 2|'+
                  '6789012 3';
  end;

  // Insert row above merged block
  with InsDelTestData[44] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertRow := 0;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '        |'+
                  '12345678|'+
                  '23456789|'+
                  '345M 890|'+
                  '456  901|'+
                  '567  012|'+
                  '67890123';
  end;

  // Insert row through merged block
  with InsDelTestData[45] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertRow := 3;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 4;
    MergedColCount_After := 2;
    SollLayout := '12345678|'+
                  '23456789|'+
                  '345M 890|'+
                  '        |'+
                  '456  901|'+
                  '567  012|'+
                  '67890123';
  end;

  // Insert row below merged block
  with InsDelTestData[46] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    InsertRow := 5;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '12345678|'+
                  '23456789|'+
                  '345M 890|'+
                  '456  901|'+
                  '567  012|'+
                  '        |'+
                  '67890123';
  end;

  // Delete column before merged block
  with InsDelTestData[47] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteCol := 1;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '1345678|'+
                  '2456789|'+
                  '35M 890|'+
                  '46  901|'+
                  '57  012|'+
                  '6890123';
  end;

  // Delete column through merged block
  with InsDelTestData[48] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteCol := 4;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 1;
    SollLayout := '1234678|'+
                  '2345789|'+
                  '345M890|'+
                  '456 901|'+
                  '567 012|'+
                  '6789123';
  end;

  // Delete column behind merged block
  with InsDelTestData[49] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteCol := 7;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '1234567|'+
                  '2345678|'+
                  '345M 89|'+
                  '456  90|'+
                  '567  01|'+
                  '6789012';
  end;

  // Delete row above merged block
  with InsDelTestData[50] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteRow := 1;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '12345678|'+
                 // '23456789|'+
                  '345M 890|'+
                  '456  901|'+
                  '567  012|'+
                  '67890123';
  end;

  // Delete row through merged block
  with InsDelTestData[51] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteRow := 4;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 2;
    MergedColCount_After := 2;
    SollLayout := '12345678|'+
                  '23456789|'+
                  '345M 890|'+
                  '456  901|'+
                //  '567  012|'+
                  '67890123';
  end;

  // Delete row behind merged block
  with InsDelTestData[52] do begin
    Layout := '12345678|'+
              '23456789|'+
              '345M 890|'+               // "M" = merged block (2 cols x 3 rows)
              '456  901|'+
              '567  012|'+
              '67890123';
    DeleteRow := 5;
    MergedRowCount := 3;
    MergedColCount := 2;
    MergedRowCount_After := 3;
    MergedColCount_After := 2;
    SollLayout := '12345678|'+
                  '23456789|'+
                  '345M 890|'+
                  '456  901|'+
                  '567  012';
  end;
end;


{ TSpreadWriteRead_InsDelColRowTests }

procedure TSpreadWriteRead_InsDelColRow_Tests.SetUp;
begin
  inherited SetUp;
  InitTestData;
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow(
  ATestIndex: Integer; AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  lastCol, lastRow: Cardinal;
  r1,c1,r2,c2: Cardinal;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  L, LL: TStringList;
  s: String;
  expected: String;
  actual: String;
  expectedFormulas: array of array of String;
begin
  TempFile := GetTempFileName;

  L := TStringList.Create;
  try
    // Extract soll formulas into a 2D array in case of shared formulas
    if (InsDelTestData[ATestIndex].SharedFormulaRowCount_After > 0) or
       (InsDelTestData[ATestIndex].SharedFormulaColCount_After > 0) then
    begin
      with InsDelTestData[ATestIndex] do
        SetLength(expectedFormulas, SharedFormulaRowCount_After, SharedFormulaColCount_After);
      L.Delimiter := ';';
      L.DelimitedText := InsDelTestData[ATestIndex].SollFormula;
      LL := TStringList.Create;
      try
        LL.Delimiter := ',';
        for row := 0 to InsDelTestData[ATestIndex].SharedFormulaRowCount_After-1 do
        begin
          s := L[row];
          LL.DelimitedText := L[row];
          for col := 0 to InsDelTestData[ATestIndex].SharedFormulaColCount_After-1 do
            expectedFormulas[row, col] := trim(LL[col]);
        end;
      finally
        LL.Free;
      end;
    end;

    L.Delimiter := '|';
    L.StrictDelimiter := true;
    L.DelimitedText := InsDelTestData[ATestIndex].Layout;

    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkSheet:= MyWorkBook.AddWorksheet(InsertColRowSheet);

      // Write out cells
      for row := 0 to L.Count-1 do
      begin
        s := L[row];
        for col := 0 to Length(s)-1 do
          case s[col+1] of
            ' '     : ; // Leave cell empty
            '0'..'9': MyWorksheet.WriteNumber(row, col, StrToInt(s[col+1]));
            'F'     : MyWorksheet.WriteFormula(row, col, InsDelTestData[ATestIndex].Formula);
            'S'     : MyWorksheet.WriteSharedFormula(
                        row,
                        col,
                        row + InsDelTestData[ATestIndex].SharedFormulaRowCount - 1,
                        col + InsDelTestData[ATestIndex].SharedFormulaColCount - 1,
                        InsDelTestData[ATestIndex].Formula
                      );
            'M'     : begin
                        MyWorksheet.WriteUTF8Text(row, col, 'M');
                        MyWorksheet.MergeCells(
                          row,
                          col,
                          row + InsDelTestData[ATestIndex].MergedRowCount - 1,
                          col + InsDelTestData[ATestIndex].MergedColCount - 1
                        );
                      end;
          end;
      end;

      if InsDelTestData[ATestIndex].InsertCol >= 0 then
        MyWorksheet.InsertCol(InsDelTestData[ATestIndex].InsertCol);

      if InsDelTestData[ATestIndex].InsertRow >= 0 then
        MyWorksheet.InsertRow(InsDelTestData[ATestIndex].InsertRow);

      if InsDelTestData[ATestIndex].DeleteCol >= 0 then
        MyWorksheet.DeleteCol(InsDelTestData[ATestIndex].DeleteCol);

      if InsDelTestData[ATestIndex].DeleteRow >= 0 then
        MyWorksheet.DeleteRow(InsDelTestData[ATestIndex].DeleteRow);

      MyWorkBook.WriteToFile(TempFile, AFormat, true);
    finally
      MyWorkbook.Free;
    end;

    L.DelimitedText := InsDelTestData[ATestIndex].SollLayout;

    // Open the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas, boAutoCalc];
      MyWorkbook.ReadFromFile(TempFile, AFormat);

      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, InsertColRowSheet);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet');

      lastRow := MyWorksheet.GetLastOccupiedRowIndex;
      lastCol := MyWorksheet.GetLastOccupiedColIndex;

      for row := 0 to lastRow do
      begin
        expected := L[row];
        actual := '';
        for col := 0 to lastcol do
        begin
          MyCell := MyWorksheet.FindCell(row, col);

          if MyCell = nil then
            actual := actual + ' '
          else
            case MyCell^.ContentType of
              cctEmpty     : actual := actual + ' ';
              cctNumber    : actual := actual + IntToStr(Round(Mycell^.NumberValue));
              cctUTF8String: actual := actual + MyCell^.UTF8StringValue;
              cctError     : actual := actual + 'E';
            end;
          if HasFormula(MyCell) then
          begin
            if (InsDelTestData[ATestIndex].SharedFormulaRowCount_After > 0) or
               (InsDelTestData[ATestIndex].SharedFormulaColCount_After > 0)
            then
              CheckEquals(
                expectedFormulas[row-InsDelTestData[ATestIndex].SharedFormulaBaseRow_After,
                                 col-InsDelTestData[ATestIndex].SharedFormulaBaseCol_After],
                MyWorksheet.ReadFormulaAsString(MyCell),
                'Shared formula mismatch, cell ' + CellNotation(MyWorksheet, Row, Col)
              )
            else
              CheckEquals(
                InsDelTestData[ATestIndex].SollFormula,
                MyWorksheet.ReadFormulaAsString(MyCell),
                'Formula mismatch, cell '+CellNotation(MyWorksheet, Row, Col)
              );
          end;
          if MyWorksheet.IsMerged(MyCell) then
          begin
            MyWorksheet.FindMergedRange(MyCell, r1, c1, r2, c2);
            CheckEquals(
              InsDelTestData[ATestIndex].MergedRowCount_After,
              r2 - r1 + 1,
              'Merged row count mismatch, cell ' + CellNotation(MyWorksheet, Row, Col)
            );
            CheckEquals(
              InsDelTestData[ATestIndex].MergedColCount_After,
              c2 - c1 + 1,
              'Merged column count mismatch, cell '+CellNotation(MyWorksheet, Row, Col)
            );
          end;
        end;
        CheckEquals(expected, actual,
          'Test empty cell layout mismatch, cell '+CellNotation(MyWorksheet, Row, Col));
      end;
    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

  finally
    L.Free;
  end;
end;


{------------------------------------------------------------------------------}
{                                 BIFF8 tests                                  }
{------------------------------------------------------------------------------}

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_0_BIFF8;
// insert a column before the first one
begin
  TestWriteRead_InsDelColRow(0, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_1_BIFF8;
// insert a column before column 2
begin
  TestWriteRead_InsDelColRow(1, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_2_BIFF8;
// insert a column before the last one
begin
  TestWriteRead_InsDelColRow(2, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_3_BIFF8;
// delete column 0
begin
  TestWriteRead_InsDelColRow(3, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_4_BIFF8;
// delete column 2
begin
  TestWriteRead_InsDelColRow(4, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_5_BIFF8;
// delete last column
begin
  TestWriteRead_InsDelColRow(5, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_6_BIFF8;
// insert row before first one
begin
  TestWriteRead_InsDelColRow(6, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_7_BIFF8;
// insert row before #2
begin
  TestWriteRead_InsDelColRow(7, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_8_BIFF8;
// insert row before last one
begin
  TestWriteRead_InsDelColRow(8, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_9_BIFF8;
// delete first row
begin
  TestWriteRead_InsDelColRow(9, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_10_BIFF8;
// delete row #2
begin
  TestWriteRead_InsDelColRow(10, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_11_BIFF8;
// delete last row
begin
  TestWriteRead_InsDelColRow(11, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_12_BIFF8;
// insert column before formula cell
begin
  TestWriteRead_InsDelColRow(12, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_13_BIFF8;
// insert column after formula cell
begin
  TestWriteRead_InsDelColRow(13, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_14_BIFF8;
// insert row before formula cell
begin
  TestWriteRead_InsDelColRow(14, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_15_BIFF8;
// insert row after formula cell
begin
  TestWriteRead_InsDelColRow(15, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_16_BIFF8;
// delete column before formula cell
begin
  TestWriteRead_InsDelColRow(16, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_17_BIFF8;
// delete column after formula cell
begin
  TestWriteRead_InsDelColRow(17, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_18_BIFF8;
// delete column containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(18, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_19_BIFF8;
// delete row before formula cell
begin
  TestWriteRead_InsDelColRow(19, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_20_BIFF8;
// delete row after formula cell
begin
  TestWriteRead_InsDelColRow(20, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_21_BIFF8;
// delete row containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(21, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_22_BIFF8;
// no insert/delete; just test shared formula
begin
  TestWriteRead_InsDelColRow(22, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_23_BIFF8;
// insert column before any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(23, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_24_BIFF8;
// insert column after any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(24, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_25_BIFF8;
// insert column through cells addressed by shared formula
begin
  TestWriteRead_InsDelColRow(25, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_26_BIFF8;
// insert column through shared formula block
begin
  TestWriteRead_InsDelColRow(26, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_27_BIFF8;
// insert column behind shared formula block
begin
  TestWriteRead_InsDelColRow(27, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_28_BIFF8;
// insert row before anything affected by the shared formula
begin
  TestWriteRead_InsDelColRow(28, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_29_BIFF8;
// insert row after shared formula cells
begin
  TestWriteRead_InsDelColRow(29, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_30_BIFF8;
// insert row through shared formula block
begin
  TestWriteRead_InsDelColRow(30, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_31_BIFF8;
// insert row below shared formula block
begin
  TestWriteRead_InsDelColRow(31, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_32_BIFF8;
// delete column before any cell referenced by the shared formula(s)
begin
  TestWriteRead_InsDelColRow(32, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_33_BIFF8;
// delete column after any cell referenced by the shared formula(s)
begin
  TestWriteRead_InsDelColRow(33, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_34_BIFF8;
// delete column in shared formula block, but no formula cell affected
begin
  TestWriteRead_InsDelColRow(34, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_35_BIFF8;
// delete column used in shared formula
begin
  TestWriteRead_InsDelColRow(35, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_36_BIFF8;
// delete row above shared formula and the used cells
begin
  TestWriteRead_InsDelColRow(36, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_37_BIFF8;
// delete row below shared formula and the used cells
begin
  TestWriteRead_InsDelColRow(37, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_38_BIFF8;
// delete row through shared formula block
begin
  TestWriteRead_InsDelColRow(38, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_39_BIFF8;
// delete row with cell used in shared formula block
begin
  TestWriteRead_InsDelColRow(39, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_40_BIFF8;
// no insert/delete, just test merged cell block
begin
  TestWriteRead_InsDelColRow(40, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_41_BIFF8;
// insert column before merged block
begin
  TestWriteRead_InsDelColRow(41, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_42_BIFF8;
// insert column through merged block
begin
  TestWriteRead_InsDelColRow(42, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_43_BIFF8;
// insert column behind merged block
begin
  TestWriteRead_InsDelColRow(43, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_44_BIFF8;
// insert row above merged block
begin
  TestWriteRead_InsDelColRow(44, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_45_BIFF8;
// insert row through merged block
begin
  TestWriteRead_InsDelColRow(45, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_46_BIFF8;
// insert row below merged block
begin
  TestWriteRead_InsDelColRow(46, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_47_BIFF8;
// delete column before merged block
begin
  TestWriteRead_InsDelColRow(47, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_48_BIFF8;
// delete column through merged block
begin
  TestWriteRead_InsDelColRow(48, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_49_BIFF8;
// delete column behind merged block
begin
  TestWriteRead_InsDelColRow(49, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_50_BIFF8;
// delete row above merged block
begin
  TestWriteRead_InsDelColRow(50, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_51_BIFF8;
// delete row through merged block
begin
  TestWriteRead_InsDelColRow(51, sfExcel8);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_52_BIFF8;
// delete row below merged block
begin
  TestWriteRead_InsDelColRow(52, sfExcel8);
end;


{ -----------------------------------------------------------------------------}
{                              OOXML Tests                                     }
{ -----------------------------------------------------------------------------}

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_0_OOXML;
// insert a column before the first one
begin
  TestWriteRead_InsDelColRow(0, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_1_OOXML;
// insert a column before column 2
begin
  TestWriteRead_InsDelColRow(1, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_2_OOXML;
// insert a column before the last one
begin
  TestWriteRead_InsDelColRow(2, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_3_OOXML;
// delete column 0
begin
  TestWriteRead_InsDelColRow(3, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_4_OOXML;
// delete column 2
begin
  TestWriteRead_InsDelColRow(4, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_5_OOXML;
// delete last column
begin
  TestWriteRead_InsDelColRow(5, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_6_OOXML;
// insert row before first one
begin
  TestWriteRead_InsDelColRow(6, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_7_OOXML;
// insert row before #2
begin
  TestWriteRead_InsDelColRow(7, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_8_OOXML;
// insert row before last one
begin
  TestWriteRead_InsDelColRow(8, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_9_OOXML;
// delete first row
begin
  TestWriteRead_InsDelColRow(9, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_10_OOXML;
// delete row #2
begin
  TestWriteRead_InsDelColRow(10, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_11_OOXML;
// delete last row
begin
  TestWriteRead_InsDelColRow(11, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_12_OOXML;
// insert column before formula cell
begin
  TestWriteRead_InsDelColRow(12, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_13_OOXML;
// insert column after formula cell
begin
  TestWriteRead_InsDelColRow(13, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_14_OOXML;
// insert row before formula cell
begin
  TestWriteRead_InsDelColRow(14, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_15_OOXML;
// insert row after formula cell
begin
  TestWriteRead_InsDelColRow(15, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_16_OOXML;
// delete column before formula cell
begin
  TestWriteRead_InsDelColRow(16, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_17_OOXML;
// delete column after formula cell
begin
  TestWriteRead_InsDelColRow(17, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_18_OOXML;
// delete column containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(18, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_19_OOXML;
// delete row before formula cell
begin
  TestWriteRead_InsDelColRow(19, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_20_OOXML;
// delete row after formula cell
begin
  TestWriteRead_InsDelColRow(20, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_21_OOXML;
// delete row containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(21, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_22_OOXML;
// no insert/delete; just test shared formula
begin
  TestWriteRead_InsDelColRow(22, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_23_OOXML;
// insert column before any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(23, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_24_OOXML;
// insert column after any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(24, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_25_OOXML;
// insert column through cells addressed by shared formula
begin
  TestWriteRead_InsDelColRow(25, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_26_OOXML;
// insert column through shared formula block
begin
  TestWriteRead_InsDelColRow(26, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_27_OOXML;
// insert column behind shared formula block
begin
  TestWriteRead_InsDelColRow(27, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_28_OOXML;
// insert row before anything affected by the shared formula
begin
  TestWriteRead_InsDelColRow(28, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_29_OOXML;
// insert row after shared formula cells
begin
  TestWriteRead_InsDelColRow(29, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_30_OOXML;
// insert row through shared formula block
begin
  TestWriteRead_InsDelColRow(30, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_31_OOXML;
// insert row below shared formula block
begin
  TestWriteRead_InsDelColRow(31, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_32_OOXML;
// delete column before any cell referenced by the shared formula(s)
begin
  TestWriteRead_InsDelColRow(32, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_33_OOXML;
// delete column after any cell referenced by the shared formula(s)
begin
  TestWriteRead_InsDelColRow(33, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_34_OOXML;
// delete column in shared formula block, but no formula cell affected
begin
  TestWriteRead_InsDelColRow(34, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_35_OOXML;
// delete column used in shared formula
begin
  TestWriteRead_InsDelColRow(35, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_36_OOXML;
// delete row above shared formula and the used cells
begin
  TestWriteRead_InsDelColRow(36, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_37_OOXML;
// delete row below shared formula and the used cells
begin
  TestWriteRead_InsDelColRow(37, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_38_OOXML;
// delete row through shared formula block
begin
  TestWriteRead_InsDelColRow(38, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_39_OOXML;
// delete row with cell used in shared formula block
begin
  TestWriteRead_InsDelColRow(39, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_40_OOXML;
// no insert/delete, just test merged cell block
begin
  TestWriteRead_InsDelColRow(40, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_41_OOXML;
// insert column before merged block
begin
  TestWriteRead_InsDelColRow(41, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_42_OOXML;
// insert column through merged block
begin
  TestWriteRead_InsDelColRow(42, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_43_OOXML;
// insert column behind merged block
begin
  TestWriteRead_InsDelColRow(43, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_44_OOXML;
// insert row above merged block
begin
  TestWriteRead_InsDelColRow(44, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_45_OOXML;
// insert row through merged block
begin
  TestWriteRead_InsDelColRow(45, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_46_OOXML;
// insert row below merged block
begin
  TestWriteRead_InsDelColRow(46, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_47_OOXML;
// delete column before merged block
begin
  TestWriteRead_InsDelColRow(47, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_48_OOXML;
// delete column through merged block
begin
  TestWriteRead_InsDelColRow(48, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_49_OOXML;
// delete column behind merged block
begin
  TestWriteRead_InsDelColRow(49, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_50_OOXML;
// delete row above merged block
begin
  TestWriteRead_InsDelColRow(50, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_51_OOXML;
// delete row through merged block
begin
  TestWriteRead_InsDelColRow(51, sfOOXML);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_52_OOXML;
// delete row below merged block
begin
  TestWriteRead_InsDelColRow(52, sfOOXML);
end;


{ -----------------------------------------------------------------------------}
{                            OpenDocument Tests                                }
{ -----------------------------------------------------------------------------}

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_0_ODS;
// insert a column before the first one
begin
  TestWriteRead_InsDelColRow(0, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_1_ODS;
// insert a column before column 2
begin
  TestWriteRead_InsDelColRow(1, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_2_ODS;
// insert a column before the last one
begin
  TestWriteRead_InsDelColRow(2, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_3_ODS;
// delete column 0
begin
  TestWriteRead_InsDelColRow(3, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_4_ODS;
// delete column 2
begin
  TestWriteRead_InsDelColRow(4, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_5_ODS;
// delete last column
begin
  TestWriteRead_InsDelColRow(5, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_6_ODS;
// insert row before first one
begin
  TestWriteRead_InsDelColRow(6, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_7_ODS;
// insert row before #2
begin
  TestWriteRead_InsDelColRow(7, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_8_ODS;
// insert row before last one
begin
  TestWriteRead_InsDelColRow(8, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_9_ODS;
// delete first row
begin
  TestWriteRead_InsDelColRow(9, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_10_ODS;
// delete row #2
begin
  TestWriteRead_InsDelColRow(10, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_11_ODS;
// delete last row
begin
  TestWriteRead_InsDelColRow(11, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_12_ODS;
// insert column before formula cell
begin
  TestWriteRead_InsDelColRow(12, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_13_ODS;
// insert column after formula cell
begin
  TestWriteRead_InsDelColRow(13, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_14_ODS;
// insert row before formula cell
begin
  TestWriteRead_InsDelColRow(14, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_15_ODS;
// insert row after formula cell
begin
  TestWriteRead_InsDelColRow(15, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_16_ODS;
// delete column before formula cell
begin
  TestWriteRead_InsDelColRow(16, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_17_ODS;
// delete column after formula cell
begin
  TestWriteRead_InsDelColRow(17, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_18_ODS;
// delete column containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(18, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_19_ODS;
// delete row before formula cell
begin
  TestWriteRead_InsDelColRow(19, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_20_ODS;
// delete row after formula cell
begin
  TestWriteRead_InsDelColRow(20, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_21_ODS;
// delete row containing a cell used in formula
begin
  TestWriteRead_InsDelColRow(21, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_22_ODS;
// no insert/delete; just test shared formula
begin
  TestWriteRead_InsDelColRow(22, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_23_ODS;
// insert column before any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(23, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_24_ODS;
// insert column after any cell addressed by the shared formula
begin
  TestWriteRead_InsDelColRow(24, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_25_ODS;
// insert column through cells addressed by shared formula
begin
  TestWriteRead_InsDelColRow(25, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_26_ODS;
// insert column through shared formula block
begin
  TestWriteRead_InsDelColRow(26, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_27_ODS;
// insert column behind shared formula block
begin
  TestWriteRead_InsDelColRow(27, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_28_ODS;
// insert row before anything affected by the shared formula
begin
  TestWriteRead_InsDelColRow(28, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_29_ODS;
// insert row after shared formula cells
begin
  TestWriteRead_InsDelColRow(29, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_30_ODS;
// insert row through shared formula block
begin
  TestWriteRead_InsDelColRow(30, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_31_ODS;
// insert row below shared formula block
begin
  TestWriteRead_InsDelColRow(31, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_32_ODS;
// delete column before any cell referenced by the shared formula(s)
begin
  TestWriteRead_InsDelColRow(32, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_33_ODS;
// delete column after any cell referenced by the shared formula(s)
begin
  TestWriteRead_InsDelColRow(33, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_34_ODS;
// delete column in shared formula block, but no formula cell affected
begin
  TestWriteRead_InsDelColRow(34, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_35_ODS;
// delete column used in shared formula
begin
  TestWriteRead_InsDelColRow(35, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_36_ODS;
// delete row above shared formula and the used cells
begin
  TestWriteRead_InsDelColRow(36, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_37_ODS;
// delete row below shared formula and the used cells
begin
  TestWriteRead_InsDelColRow(37, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_38_ODS;
// delete row through shared formula block
begin
  TestWriteRead_InsDelColRow(38, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_39_ODS;
// delete row with cell used in shared formula block
begin
  TestWriteRead_InsDelColRow(39, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_40_ODS;
// no insert/delete, just test merged cell block
begin
  TestWriteRead_InsDelColRow(40, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_41_ODS;
// insert column before merged block
begin
  TestWriteRead_InsDelColRow(41, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_42_ODS;
// insert column through merged block
begin
  TestWriteRead_InsDelColRow(42, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_43_ODS;
// insert column behind merged block
begin
  TestWriteRead_InsDelColRow(43, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_44_ODS;
// insert row above merged block
begin
  TestWriteRead_InsDelColRow(44, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_45_ODS;
// insert row through merged block
begin
  TestWriteRead_InsDelColRow(45, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_46_ODS;
// insert row below merged block
begin
  TestWriteRead_InsDelColRow(46, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_47_ODS;
// delete column before merged block
begin
  TestWriteRead_InsDelColRow(47, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_48_ODS;
// delete column through merged block
begin
  TestWriteRead_InsDelColRow(48, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_49_ODS;
// delete column behind merged block
begin
  TestWriteRead_InsDelColRow(49, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_50_ODS;
// delete row above merged block
begin
  TestWriteRead_InsDelColRow(50, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_51_ODS;
// delete row through merged block
begin
  TestWriteRead_InsDelColRow(51, sfOpenDocument);
end;

procedure TSpreadWriteRead_InsDelColRow_Tests.TestWriteRead_InsDelColRow_52_ODS;
// delete row below merged block
begin
  TestWriteRead_InsDelColRow(52, sfOpenDocument);
end;



initialization
  RegisterTest(TSpreadWriteRead_InsDelColRow_Tests);
  InitTestData;

end.

