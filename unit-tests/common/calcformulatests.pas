unit calcformulatests;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpstypes, fpspreadsheet, fpsexprparser;

type

  TCalcFormulaTests = class(TTestCase)
  private
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    FOtherWorksheet: TsWorksheet;
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
  end;

  TCalcDateTimeFormulaTests = class(TCalcFormulaTests)
  published
    procedure Test_DATE;
//    procedure Test_DATEDIF;  to be written
//    procedure Test_DATEVALUE;  to be written
//    procedure Test_DAY;        to be written
//    procedure Test_HOUR;       to be written
//    procedure Test_MINUTE;     to be written
//    procedure Test_MONTH;      to be written
//    procedure Test_NOW;        to be written
//    procedure Test_SECOND;     to be written
    procedure Test_TIME;
//    procedure Test_TIMEVALUE;  to be written
//    procedure Test_TODAY;      to be written
//    procedure Test_WEEKDAY;    to be written
//    procedure Test_YEAR;       to be written
  end;

  TCalcInfoFormulaTests = class(TCalcFormulaTests)
  published
    procedure Test_ERRORTYPE;
    procedure Test_IFERROR;
    procedure Test_ISBLANK;
    procedure Test_ISERR;
    procedure Test_ISERROR;
    procedure Test_ISLOGICAL;
    procedure Test_ISNA;
    procedure Test_ISNONTEXT;
    procedure Test_ISNUMBER;
    procedure Test_ISREF;
    procedure Test_ISTEXT;
  end;

  TCalcLogicalFormulaTests = class(TCalcFormulaTests)
  published
    procedure Test_AND;
//    procedure Test_FALSE;  to be written
    procedure Test_IF;
//    procedure Test_XLFN.IFS;  to be written
    procedure Test_NOT;
    procedure Test_OR;
//    procedure Test_TRUE;  to be written
  end;

  TCalcLookupFormulaTests = class(TCalcFormulaTests)
  published
    procedure Test_ADDRESS;
    procedure Test_COLUMN;
    // procedure Test_HYPERLINK  -- to be written
    procedure Test_INDEX_1;
    procedure Test_INDEX_2;
    procedure Test_INDIRECT;
    procedure Test_MATCH;
    procedure Test_ROW;
  end;

  TCalcMathFormulaTests = class(TCalcFormulaTests)
  published
    procedure Test_ABS;
    procedure Test_ACOS;
//    procedure Test_ACOSH;     to be written
    procedure Test_ASIN;
//    procedure Test_ASINH;     to be written
    procedure Test_ATAN;
//    procedure Test_ATANH;     to be written
    procedure Test_CEILING;
//    procedure Test_COS;       to be written
//    procedure Test_COSH;      to be written
//    procedure Test_DEGREES;   to be written
    procedure Test_EVEN;
//    procedure Test_EXP;       to be written
//    procedure Test_FACT;      to be written
    procedure Test_FLOOR;
//    procedure Test_INT;       to be written
//    procedure Test_LN;        to be written
    procedure Test_LOG;
    procedure Test_LOG10;
//    procedure Test_MOD;       to be written
    procedure Test_ODD;
//    procedure Test_PI;        to be written
    procedure Test_POWER;
    procedure Test_RADIANS;
//    procedure Test_RAND;      to be written
    procedure Test_ROUND;
//    procedure Test_ROUNDDOWN; to be written
//    procedure Test_ROUNDUP;   to be written
//    procedure Test_SIGN;      to be written
//    procedure Test_SIN;       to be written
//    procedure Test_SINH;      to be written
    procedure Test_SQRT;
//    procedure Test_TAN;       to be written
//    procedure Test_TANH;      to be written
  end;

  TCalcStatsFormulaTests = class(TCalcFormulaTests)
  published
    procedure Test_AVEDEV;
    procedure Test_AVERAGE;
    procedure Test_AVERAGEIF;
    // procedure Test_AVERAGEIFS;   to be written
    procedure Test_COUNT;
    procedure Test_COUNTA;
    procedure Test_COUNTBLANK;
    procedure Test_COUNTIF;
//    procedure Test_COUNITFS;  to be written
    procedure Test_MAX;
    procedure Test_MIN;
    procedure Test_PRODUCT;
    procedure Test_STDEV;
    procedure Test_STDEVP;
    procedure Test_SUM;
    procedure Test_SUMIF;
//    procedure Test_SUMIFS;  to be written
    procedure Test_SUMSQ;
    procedure Test_VAR;
    procedure Test_VARP;
  end;

  TCalcTextFormulaTests = class(TCalcFormulaTests)
  published
//    procedure Test_CHAR;       to be written
//    procedure Test_CODE;       to be written
    procedure Test_CONCATENATE;
    procedure Test_EXACT;
//    procedure Test_LEFT;       to be written
    procedure Test_LEN;
    procedure Test_LOWER;
//    procedure Test_MID;        to be written
//    procedure Test_REPLACE;    to be written
//    procedure Test_REPT;       to be written
//    procedure Test_RIGHT;      to be written
//    procedure Test_SUBSTITUTE; to be written
//    procedure Test_TEXT;       to be written
//    procedure Test_TRIM;       to be written
    procedure Test_UPPER;
//    procedure Test_VALUE;      to be written
  end;

implementation

procedure TCalcFormulaTests.Setup;
begin
  FWorkbook := TsWorkbook.Create;
  FWorksheet := FWorkbook.AddWorksheet('Sheet1');
  FOtherWorksheet := FWorkbook.AddWorksheet('Sheet2');
end;

procedure TCalcFormulaTests.TearDown;
begin
  FWorkbook.Free;
end;

// *** Date/time formula tests
{$I testcases_calcdatetimeformulas.inc}

// *** Information formula tests
{$I testcases_calcinfoformulas.inc}

// *** Logical formula tests
{$I testcases_calclogicalformulas.inc}

// *** Lookup formula tests
{$I testcases_calclookupformulas.inc}

// *** Math formula tests
{$I testcases_calcmathformulas.inc}

// *** Statistical formula tests
{$I testcases_calcstatsformulas.inc}

// *** Text formula tests
{$I testcases_calctextformulas.inc}

initialization
  RegisterTest('TCalcFormulaTests', TCalcDateTimeFormulaTests);
  RegisterTest('TCalcFormulaTests', TCalcInfoFormulaTests);
  RegisterTest('TCalcFormulaTests', TCalcLogicalFormulaTests);
  RegisterTest('TCalcFormulaTests', TCalcLookupFormulaTests);
  RegisterTest('TCalcFormulaTests', TCalcMathFormulaTests);
  RegisterTest('TCalcFormulaTests', TCalcStatsFormulaTests);
  RegisterTest('TCalcFormulaTests', TCalcTextFormulaTests);
end.

