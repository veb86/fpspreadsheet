{-------------------------------------------------------------------------------
   Tests for some dedicated math routines which are specific to spreadsheets.
-------------------------------------------------------------------------------}

unit mathtests;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  {$IFDEF Unix}
  //required for formatsettings
  clocale,
  {$ENDIF}
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry, testsutility,
  fpstypes, fpspreadsheet, fpsutils;

type
  { TSpreadMathTests }
  //Write to xls/xml file and read back
  TSpreadMathTests = class(TTestCase)
  private
  protected
    procedure TestRound(InputValue: Double; Expected: Integer);

  published
    // Test whether "round" avoids Banker's rounding
    procedure TestRound_plus15;
    procedure Testround_minus15;
    procedure TestRound_plus25;
    procedure TestRound_minus25;

  end;

implementation

{ TSpreadMathTests }

procedure TSpreadMathTests.TestRound(InputValue: Double; Expected: Integer);
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  readValue: String;
begin
  book := TsWorkbook.Create;
  try
    sheet := book.AddWorksheet('Math');
    sheet.WriteNumber(1, 1, InputValue, nfFixed, 0);
    readValue := sheet.ReadAsText(1, 1);

    CheckEquals(Expected, StrToInt(readValue),
      'Rounding error, sheet "' + sheet.Name + '"')
  finally
    book.Free;
  end;
end;

procedure TSpreadMathTests.TestRound_plus15;
begin
  TestRound(1.5, 2);
end;

procedure TSpreadMathTests.TestRound_minus15;
begin
  Testround(-1.5, -2);
end;

procedure TSpreadMathTests.TestRound_plus25;
begin
  TestRound(2.5, 3);
end;

procedure TSpreadMathTests.Testround_minus25;
begin
  TestRound(-2.5, -3);
end;


initialization
  RegisterTest(TSpreadMathTests);

end.

