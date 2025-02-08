{------------------------------------------------------------------------------}
{  Standard built-in formula support                                           }
{------------------------------------------------------------------------------}

unit fpsfunc;

{$mode objfpc}{$H+}
{$WARN 6058 off : Call to subroutine "$1" marked as inline is not inlined}
interface

uses 
  Classes, SysUtils, fpstypes;

function CompareStringWithWildcards(AString1, AString2: String): Boolean;
procedure RegisterStdBuiltins(AManager: TComponent);


implementation

uses
  Math, lazutf8, StrUtils, DateUtils,
  xlsconst, {%H-}fpsPatches, fpsUtils,
  fpsnumformat, fpspreadsheet, fpsexprparser;

const
  EPS = 1E-12;

{ AString1 is the string with wildcards. }
function CompareStringWithWildCards(AString1, AString2: String): Boolean;
var
  P1, P2, P10, P20: PAnsiChar;
  n1, n2: Integer;
begin
  if (AString1 = '') and (AString2 = '') then
  begin
    Result := true;
    exit;
  end;
  if ((AString1 <> '') and (AString2 = '')) or ((AString1 = '') and (AString2 <> '')) then
  begin
    Result := false;
    exit;
  end;

  n1 := Length(AString1);
  n2 := Length(AString2);
  P10 := PAnsiChar(AString1);
  P20 := PAnsiChar(AString2);
  P1 := P10;
  P2 := P20;
  while P1 < P10 + n1 do begin
    if P1^ = '*' then
    begin
      Result := True;
      exit;
    end else
    if (P1^ = '?') or (P1^ = P2^) then begin
      inc(P1);
      if P2 < P20 + n2 then
        inc(P2)
      else
      begin
        Result := false;
        exit;
      end;
    end else
    if (P1^ <> P2^) then
    begin
      Result := false;
      exit;
    end;
  end;
  Result := true;
end;

type
  TsFuncType = (ftCountIF, ftCountIFS, ftSumIF, ftSUMIFS, ftAverageIF, ftAverageIFS);
  TsCompareType = (ctNumber, ctString, ctBoolean, ctEmpty, ctError);

  { Helper class for calculating COUNTIF(S) or SUMIF(S) or AVERAGEIF(S) formulas.

    Parameters are defined in the constructor:
    - Args: array of TsExpressionParameters, provided by the expression parser.
    - AValueRangeIndex: Index in Args defining the parameters for the range,
      in which values are added for SUMIF(S), AVERAGEIF(S).
    - ACriteriaRangeIndex: Index in Args defining the first range parameter
      for comparing with the criterial.
    - ACriteriaIndex: Index in Args defining the comparison expression, e.g. '>10'
    - AFuncType: defines the function to be calculated.

    In COUNTIFS, SUMIFS, AVERAGEIFS the criteria range and criteria parameters
    can be repeated (up to 127, by Excel specification, not checked here).
  }
  TsFuncComparer = class
  private
    FArgs: TsExprParameterArray;
    FValueRangeIndex: Integer;
    FCriteriaRangeIndex: Integer;
    FCriteriaIndex: Integer;
    FFuncType: TsFuncType;
    FCompareOperation: TsCompareOperation;
    FCompareType: TsCompareType;
    FCompareNumber: Double;
    FCompareBoolean: Boolean;
    FCompareError: TsErrorValue;
    FCompareString: String;
    FFormatSettings: TFormatSettings;
    FError: TsErrorValue;
  protected
    function CompareArg(ArgIndex: Integer): Boolean;
    function CompareBoolean(AValue: Boolean): Boolean;
    function CompareCell(ASheet: TsBasicWorksheet; ARow, ACol: Integer): Boolean;
    function CompareEmpty(AEmpty: Boolean): Boolean;
    function CompareError(AError: TsErrorValue): Boolean;
    function CompareNumber(ANumber: Double): Boolean;
    function CompareString(AString: String): Boolean;
    function GetArgValue(ArgIndex: Integer): Double;
    function GetCellValue(ASheet: TsBasicWorksheet; ARow, ACol: Integer): Double;
    procedure GetCompareParams(ArgIndex: Integer);
    procedure GetRangeLimits(ArgIndex: Integer; out ARow1, ACol1, ARow2, ACol2: Integer);
    function GetWorkbook: TsBasicWorkbook;
    function GetWorksheet(ArgIndex: Integer): TsBasicWorksheet;
    function SameRangeSize(ARange1, ARange2: TsCellRange3D): Boolean;
    function ValidParams(var AValue: TsExpressionResult): Boolean;

  public
    constructor Create(AArgs: TsExprParameterArray;
      AValueRangeIndex, ACriteriaRangeIndex, ACriteriaIndex: Integer;
      AFuncType: TsFuncType);
    function Execute: TsExpressionResult;
  end;

constructor TsFuncComparer.Create(AArgs: TsExprParameterArray;
  AValueRangeIndex, ACriteriaRangeIndex, ACriteriaIndex: Integer;
  AFuncType: TsFuncType);
begin
  FArgs := AArgs;
  FValueRangeIndex := AValueRangeIndex;
  FCriteriaRangeIndex := ACriteriaRangeIndex;
  FCriteriaIndex := ACriteriaIndex;
  FFuncType := AFuncType;
end;

function TsFuncComparer.CompareBoolean(AValue: Boolean): Boolean;
var
  val: Double;
begin
  Result := false;
  case FCompareOperation of
    coEqual        : if AValue = FCompareBoolean then Result := true;
    coLess         : if AValue < FCompareBoolean then Result := true;
    coGreater      : if AValue > FCompareBoolean then Result := true;
    coLessEqual    : if AValue <= FCompareBoolean then Result := true;
    coGreaterEqual : if AValue >= FCompareBoolean then Result := true;
    coNotEqual     : if AValue <> FCompareBoolean then Result := true;
  end;
end;

function TsFuncComparer.CompareNumber(ANumber: Double): Boolean;
begin
  Result := false;
  case FCompareOperation of
    coEqual        : if ANumber = FCompareNumber then Result := true;
    coLess         : if ANumber < FCompareNumber then Result := true;
    coGreater      : if ANumber > FCompareNumber then Result := true;
    coLessEqual    : if ANumber <= FCompareNumber then Result := true;
    coGreaterEqual : if ANumber >= FCompareNumber then Result := true;
    coNotEqual     : if ANumber <> FCompareNumber then Result := true;
  end;
end;

function TsFuncComparer.CompareString(AString: String): Boolean;
begin
  Result := false;
  AString := UTF8Lowercase(AString);
  case FCompareOperation of
    coEqual        : Result := CompareStringWithWildCards(FCompareString, AString);
    coLess         : if AString < FCompareString then Result := true;
    coGreater      : if AString > FCompareString then Result := true;
    coLessEqual    : if AString <= FCompareString then Result := true;
    coGreaterEqual : if AString >= FCompareString then Result := true;
    coNotEqual     : Result := not CompareStringWithWildCards(FCompareString, AString);
  end;
end;

function TsFuncComparer.CompareEmpty(AEmpty: Boolean): Boolean;
begin
  Result := false;
  case FCompareOperation of
    coEqual        : if AEmpty then Result := true;
    coNotEqual     : if not AEmpty then Result := true;
  end;
end;

function TsFuncComparer.CompareError(AError: TsErrorValue): Boolean;
begin
  Result := false;
  case FCompareOperation of
    coEqual : Result := AError = FCompareError;
    coNotEqual: Result := AError <> FCompareError;
  end;
end;

function TsFuncComparer.CompareArg(ArgIndex: Integer): Boolean;
begin
  case FCompareType of
    ctNumber : Result := CompareNumber(ArgToFloat(FArgs[ArgIndex]));
    ctString : Result := CompareString(ArgToString(FArgs[ArgIndex]));
    ctBoolean: Result := CompareBoolean(ArgToBoolean(FArgs[ArgIndex], true));
    ctEmpty  : Result := CompareEmpty((ArgToString(FArgs[ArgIndex])) = '');
    ctError  : Result := CompareError(ArgToError(FArgs[ArgIndex]));
    else       Result := false;
  end
end;

function TsFuncComparer.CompareCell(ASheet: TsBasicWorksheet; ARow, ACol: Integer): Boolean;
var
  cell: PCell;
  value: Double;
begin
  Result := false;
  cell := TsWorksheet(ASheet).FindCell(ARow, ACol);
  case FCompareType of
    ctNumber:
      if (FCompareOperation = coNotEqual) and
         ((cell = nil) or (not (cell^.ContentType in [cctNumber, cctDateTime, cctBool])))
      then
        Result := true
      else
      if cell <> nil then
      begin
        case cell^.ContentType of
          cctNumber:
            Result := CompareNumber(cell^.NumberValue);
          cctDateTime:
            Result := CompareNumber(cell^.DateTimeValue);
          cctBool:
            if FFuncType <> ftCountIF then
              Result := CompareBoolean(cell^.Boolvalue);
          cctUTF8String:
            begin
              if TryStrToFloat(cell^.UTF8StringValue, value) then
                Result := CompareNumber(value);
            end;
          cctError:
            Result := false;
        end;
      end;
    ctString:
      if (FCompareOperation = coNotEqual) and ((cell = nil) or (cell^.ContentType <> cctUTF8String)) then
        Result := true
      else
      if (cell <> nil) then
      begin
        if (cell^.ContentType = cctUTF8String) then
          Result := CompareString(cell^.Utf8StringValue);
      end;
    ctBoolean:
      if (FCompareOperation = coNotEqual) and ((cell = nil) or (cell^.ContentType <> cctBool)) then
        Result := true
      else
      if (cell <> nil) and (cell^.ContentType = cctBool) then
        Result := CompareBoolean(cell^.BoolValue);
    ctEmpty:
      Result := CompareEmpty((cell = nil) or ((cell <> nil) and (cell^.ContentType = cctEmpty)));
    ctError:
      if (FCompareOperation = coNotEqual) and ((cell = nil) or (cell^.ContentType <> cctError)) then
        Result := true
      else
      if (cell <> nil) and (cell^.ContentType = cctError) then
        Result := CompareError(cell^.ErrorValue);
  end;
end;

{ Main method of the class: Evaluates the parameters given in the constructor and
  returns the result to be used by the formula engine. }
function TsFuncComparer.Execute: TsExpressionResult;
var
  r, r1, r2, c, c1, c2, rIdx, cIdx: Integer;
  critIdx: Integer;
  critRangeIdx: Integer;
  critSheet: TsBasicWorksheet;
  valueSheet: TsBasicWorksheet;
  matches: Boolean;
  count: Integer;
  val, sum: Double;
begin
  Result := ErrorResult(errArgError);
  FError := errOK;

  if not ValidParams(Result) then
    exit;

  // Get format settings for string-to-float or string-to-datetime conversions
  FFormatSettings := GetWorkbook.FormatSettings;

  // Get worksheet with the values to be counted, added, ...
  valueSheet := GetWorksheet(FValueRangeIndex);

  // Initialize result variables
  count := 0;
  sum := 0.0;

  // Iterate over all value range cells
  GetRangeLimits(FValueRangeIndex, r1, c1, r2, c2);
  for r := r1 to r2 do
  begin
    for c := c1 to c2 do
    begin
      // Iterate over criteria and criteria ranges: all criteria must be fulfilled.
      critIdx := FCriteriaIndex;
      critRangeIdx := FCriteriaRangeIndex;
      matches := true;
      while matches and (critIdx < Length(FArgs)) and (critRangeIdx < Length(FArgs)) do
      begin
        // Get worksheet containing the criteria range
        critSheet := GetWorksheet(critRangeIdx);

        // Analyze the compare expression
        GetCompareParams(critIdx);

        // Empty cells cannot be checked for <=, <, >, >=  --> error
        if (FCompareType = ctEmpty) and not (FCompareOperation in [coEqual, coNotEqual]) then
        begin
          Result := ErrorResult(errArgError);
          exit;
        end;
                  (*
        // Strings cannot be added --> error
        if (FFuncType in [ftSUMIF, ftSUMIFS, ftAVERAGEIF, ftAVERAGEIFS]) and (FCompareType = ctString) then
        begin
          Result := ErrorResult(errArgError);
          exit;
        end;
        *)

        // Compare current criteria cell with criteria. All criteria are "AND"-ed.
        case FArgs[critRangeIdx].ResultType of
          rtCell:
            matches := matches and CompareArg(critRangeIdx);
          rtCellRange:
            begin
              rIdx := FArgs[critRangeIdx].ResCellRange.Row1 + r - r1;
              cIdx := FArgs[critRangeIdx].ResCellRange.Col1 + c - c1;
              matches := matches and CompareCell(critSheet, rIdx, cIdx);
            end;
        end;

        inc(critIdx, 2);
        inc(critRangeIdx, 2);
        if not matches then
          break;
      end;  // while

      if matches then
      begin
        inc(count);
        case FArgs[FValueRangeIndex].ResultType of
          rtCell: sum := sum + GetArgValue(FValueRangeIndex);
          rtCellRange:
            begin
              val := GetCellValue(valuesheet, r, c);
              if not (FFuncType in [ftCountIF, ftCountIFS]) and (FError <> errOK) then
              begin
                Result := ErrorResult(FError);
                exit;
              end;
              if IsNaN(val) then
                dec(count)
              else
                sum := sum + val;
            end;
        end;
      end;
    end;  // for c
  end;  // for r

  // Final result
  case FFuncType of
    ftCOUNTIF, ftCOUNTIFS:
      Result := IntegerResult(count);
    ftSUMIF, ftSUMIFS:
      Result := FloatResult(sum);
    ftAVERAGEIF, ftAVERAGEIFS:
      if count > 0 then
        Result := FloatResult(sum / count)
      else
        Result := ErrorResult(errDivideByZero);
  end;
end;

function TsFuncComparer.GetArgValue(ArgIndex: Integer): Double;
begin
  Result := ArgToFloat(FArgs[ArgIndex])
end;

function TsFuncComparer.GetCellValue(ASheet: TsBasicWorksheet; ARow, ACol: Integer): Double;
var
  cell: PCell;
begin
  Result := 0.0;
  if FFuncType in [ftSUMIF, ftSUMIFS, ftAVERAGEIF, ftAVERAGEIFS] then
  begin
    cell := TsWorksheet(ASheet).FindCell(ARow, ACol);
    if cell <> nil then
      case cell^.ContentType of
        cctNumber:
          Result := cell^.NumberValue;
        cctDateTime:
          Result := cell^.DateTimeValue;
        cctBool:
          Result := NaN;
        cctUTF8String:
          if not TryStrToFloat(cell^.UTF8StringValue, Result) then Result := NaN;
        cctError:
          begin
            FError := cell^.ErrorValue;
            Result := NaN;
          end;
      end;
  end;
end;

{ Analyzes the criteria argument and extracts the parameters for comparing, e.g.
  '>10' ---> FCompareOperation = coGreater, FCompareNumber = 10 }
procedure TsFuncComparer.GetCompareParams(ArgIndex: Integer);
var
  cell: PCell;
  s: String;
  n: Integer;
  x: Double;
  dt: TDateTime;
begin
  FCompareOperation := coEqual;  // Default: Check for equality

  if (FArgs[ArgIndex].ResultType = rtCell) then
  begin
    cell := ArgToCell(FArgs[ArgIndex]);
    if cell = nil then
      FCompareType := ctEmpty
    else
      case cell^.ContentType of
        cctNumber:
          begin
            FCompareNumber := cell^.NumberValue;
            FCompareType := ctNumber;
          end;
        cctDateTime:
          begin
            FCompareNumber := cell^.DateTimevalue;
            FCompareType := ctNumber;
          end;
        cctBool:
          begin
            FCompareBoolean := cell^.BoolValue;
            FCompareType := ctBoolean;
          end;
        cctUTF8String:
          begin
            FCompareString := UTF8Lowercase(cell^.UTF8StringValue);
            FCompareType := ctString;
          end;
        cctEmpty:
          begin
            FCompareType := ctEmpty;
          end;
        cctError:
          begin
            FCompareError := cell^.ErrorValue;
            FCompareType := ctError;
          end;
      end;
  end else
  begin
    s := ArgToString(FArgs[ArgIndex]);
    if (Length(s) >= 1) and (s[1] in ['=', '<', '>']) then
      s := AnalyzeCompareStr(s, FCompareOperation);
    if s = '' then
      FCompareType := ctEmpty
    else
    if (FArgs[ArgIndex].ResultType = rtError) then
    begin
      FCompareError := FArgs[ArgIndex].ResError;
      FCompareType := ctError;
    end else
    if (FArgs[ArgIndex].ResultType = rtBoolean) then
    begin
      FCompareBoolean := FArgs[ArgIndex].ResBoolean;
      FCompareType := ctBoolean;
    end else
    if TryStrToInt(s, n) then
    begin
      FCompareNumber := n;
      FCompareType := ctNumber;
    end else
    if TryStrToFloat(s, x, FFormatSettings) then
    begin
      FCompareNumber := x;
      FCompareType := ctNumber;
    end else
    if TryStrToDate(s, dt, FFormatSettings) or
       TryStrToTime(s, dt, FFormatSettings) or
       TryStrToDateTime(s, dt, FFormatSettings) then
    begin
      FCompareNumber := dt;
      FCompareType := ctNumber;
    end else
    begin
      FCompareString := UTF8Lowercase(s);
      FCompareType := ctString;
    end;
  end;
end;

procedure TsFuncComparer.GetRangeLimits(ArgIndex: Integer; out ARow1, ACol1, ARow2, ACol2: Integer);
begin
  case FArgs[ArgIndex].ResultType of
  rtCell:
    begin
      ARow1 := FArgs[ArgIndex].ResRow;   ARow2 := ARow1;
      ACol1 := FArgs[ArgIndex].ResCol;   ACol2 := ACol1;
    end;
  rtCellRange:
    begin
      ARow1 := FArgs[ArgIndex].ResCellRange.Row1;
      ARow2 := FArgs[ArgIndex].ResCellRange.Row2;
      ACol1 := FArgs[ArgIndex].ResCellRange.Col1;
      ACol2 := FArgs[ArgIndex].ResCellRange.Col2;
    end;
  end;
end;

{ Extracts the workbook from the input arguments. Casting to TsWorkbook may be
  required.
  NOTE: It is assumed that all ranges are in the same workbook !!! }
function TsFuncComparer.GetWorkbook: TsBasicWorkbook;
begin
  if FArgs[FValueRangeIndex].ResultType in [rtCell, rtCellRange] then
    Result := TsWorksheet(FArgs[FValueRangeIndex].Worksheet).Workbook
  else
    Result := nil;
end;

function TsFuncComparer.GetWorksheet(ArgIndex: Integer): TsBasicWorksheet;
var
  book: TsWorkbook;
begin
  book := TsWorkbook(GetWorkbook);
  case FArgs[ArgIndex].ResultType of
    rtCell:
      Result := book.GetWorksheetByIndex(FArgs[ArgIndex].ResSheetIndex);
    rtCellRange:
      Result := book.GetWorksheetByIndex(FArgs[ArgIndex].ResCellRange.Sheet1);
      // We do not support 3d ranges here, i.e. Sheet1 = Sheet2 !!!
    else
      Result := nil;
  end;
end;

{ Checks whether the two ranges have the same size (#rows, #cols). }
function TsFuncComparer.SameRangeSize(ARange1, ARange2: TsCellRange3D): Boolean;
begin
  Result := ((ARange1.Row2 - ARange1.Row1) = (ARange2.Row2 - ARange2.Row1)) and
            ((ARange1.Col2 - ARange1.Col1) = (ARange2.Col2 - ARange2.Col1));
end;

{ Checks consistency of the input parameters passed to the constructor. }
function TsFuncComparer.ValidParams(var AValue: TsExpressionResult): Boolean;
var
  i: Integer;
begin
  Result := false;
  if Length(FArgs) = 0 then
  begin
    AValue := IntegerResult(0);
    exit;
  end;

  if Length(FArgs) = 1 then
  begin
    if IsError(FArgs[0], AValue) then
      exit;
    AValue := ErrorResult(errArgError);
    exit;
  end;

  // Special requirement for xxxIFS operations

  // Pairs of criteria range and criteria arguments required in addition to valuerange
  case FFuncType of
    ftCOUNTIFS:
      if (Length(FArgs) mod 2) <> 0 then
      begin
        AValue := ErrorResult(errArgError);
        exit;
      end;
    ftSUMIFS, ftAVERAGEIFS:
      if (Length(FArgs) - 1) mod 2 <> 0 then
      begin
        AValue := ErrorResult(errArgError);
        exit;
      end;
  end;

  if (FFuncType in [ftCOUNTIFS, ftSUMIFS, ftAVERAGEIFS]) then
  begin
    // All ranges must have the same dimensions
    if FArgs[FValueRangeIndex].ResultType = rtCell then
    begin
      if FFuncType = ftCOUNTIFS then i := 0 else i := 1;
      while (i < Length(FArgs)) do
      begin
        if FArgs[i].ResultType <> rtCell then
        begin
          AValue := ErrorResult(errArgError);
          exit;
        end;
        inc(i);
      end;
    end else
    if FArgs[FValueRangeIndex].ResultType = rtCellRange then
    begin
      if FFuncType = ftCOUNTIFS then i := 0 else i := 1;
      while (i < Length(FArgs)) do
      begin
        if FArgs[i].ResultType <> rtCellRange then
        begin
          AValue := ErrorResult(errArgError);
          exit;
        end;
        if not SameRangeSize(FArgs[FValueRangeIndex].ResCellRange, FArgs[i].ResCellRange) then
        begin
          AValue := ErrorResult(errArgError);
          exit;
        end;
        inc(i, 2);
      end;
    end;
  end;

  // To do: SUMIF and AVERAGEIF can have different range sizes!

  Result := true;
end;




{------------------------------------------------------------------------------}
{   Builtin math functions                                                     }
{------------------------------------------------------------------------------}

procedure fpsABS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(abs(x));
end;

procedure fpsACOS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
  if InRange(x, -1, +1) then
    Result := FloatResult(arccos(x))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

procedure fpsACOSH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
  if x >= 1 then
    Result := FloatResult(arccosh(x))
  else
    Result := ErrorResult(errOverflow);
end;

procedure fpsASIN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
  if InRange(x, -1, +1) then
    Result := FloatResult(arcsin(x))
  else
    Result := ErrorResult(errOverflow);
end;

procedure fpsASINH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(arcsinh(x));
end;

procedure fpsATAN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(arctan(x));
end;

procedure fpsATANH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
  if (x > -1) and (x < +1) then
    Result := FloatResult(arctanh(x))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

{    CEILING( number, significance )

  Returns a number rounded up to a multiple of significance

  - If either argument is nonnumeric, CEILING returns the #VALUE! error value.
  - Regardless of the sign of number, a value is rounded up when adjusted
    away from zero.
  - If number is an exact multiple of significance, no rounding occurs.
  - If number and significance have different signs, CEILING returns the #NUM! error value.
}
procedure fpsCEILING(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  num, sig: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  num := ArgToFloatOrNaN(Args[0]);
  sig := ArgToFloatOrNaN(Args[1]);
  if IsNaN(num) or IsNaN(sig) then      // non-numeric --> #VALUE!
    Result := ErrorResult(errWrongType)
  else
  if num*sig < 0 then
    Result := ErrorResult(errOverflow)  // arguments must have same sign, otherwise #NUM!
  else
  if sig = 0 then
    Result := FloatResult(0.0)           // as tested in Excel...
  else
  begin
    if sig < 0 then sig := -sig;
    Result := FloatResult(ceil(num/sig)*sig);
  end;
end;

procedure fpsCOS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(cos(x));
end;

procedure fpsCOSH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(cosh(x));
end;

procedure fpsDEGREES(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(RadToDeg(x));
end;

procedure fpsEVEN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// EVEN( number )
// rounds a number up to the nearest even integer.
// If the number is negative, the number is rounded away from zero.
var
  x: TsExprFloat;
  n: Integer;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  if Args[0].ResultType in [rtCell, rtInteger, rtFloat, rtDateTime, rtEmpty] then begin
    x := ArgToFloat(Args[0]);
    if not IsNaN(x) then
    begin
      if x > 0 then
      begin
        n := Trunc(x) + 1;
        if odd(n) then inc(n);
      end else
      if x < 0 then
      begin
        n := Trunc(x) - 1;
        if odd(n) then dec(n);
      end else
        n := 0;
      Result := IntegerResult(n);
    end;
  end;
end;

procedure fpsEXP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(exp(x));
end;

procedure fpsFACT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// FACT( number )
//  returns the factorial of a number.
var
  res: TsExprFloat;
  i, n: Integer;
begin
  if IsError(Args[0], Result) then
    exit;

  if Args[0].ResultType in [rtCell, rtInteger, rtFloat, rtEmpty, rtDateTime] then
  begin
    res := 1.0;
    n := ArgToInt(Args[0]);
    if n < 0 then
      Result := ErrorResult(errOverflow)
    else
      try
        for i:=1 to n do
          res := res * i;
        Result := FloatResult(res);
      except on E: EFPSpreadsheet do
        Result := ErrorResult(errOverflow);
      end;
  end else
    Result := ErrorResult(errWrongType);
end;

{  FLOOR( number, significance )
  returns a number rounded down to a multiple of significance

  - If either argument is nonnumeric, FLOOR returns the #VALUE! error value.
  - If number is positive and significance is negative, FLOOR returns the
    #NUM! error value.
  - If the sign of number is positive, a value is rounded down and adjusted
    toward zero.
  - If the sign of number is negative, a value is rounded down and adjusted
    away from zero.
  - If number is an exact multiple of significance, no rounding occurs.
}
procedure fpsFLOOR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  num, sig: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  num := ArgToFloatOrNaN(Args[0]);
  sig := ArgToFloatOrNaN(Args[1]);
  if IsNaN(num) or IsNaN(sig) then      // non-numeric --> #VALUE!
    Result := ErrorResult(errWrongType)
  else
  if sig = 0 then
    Result := FloatResult(0.0)         // as tested in Excel...
  else
  if (num > 0) and (sig < 0) then
    Result := ErrorResult(errOverflow)
  else
  begin
    Result := FloatResult(floor(num/sig)*sig);
  end;
end;

procedure fpsINT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(floor(x));
end;

procedure fpsLN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
  if x > 0 then
    Result := FloatResult(ln(x))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

procedure fpsLOG(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LOG( number [, base] )  -  base is 10 if omitted.
var
  x: TsExprFloat;
  base: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then begin
    Result := ErrorResult(errWrongType);
    exit;
  end;

  if x <= 0 then begin
    Result := ErrorResult(errOverflow);  // #NUM!
    exit;
  end;

  if (Length(Args) = 2) then
  begin
    if (Args[1].ResultType = rtMissingArg) then
    begin
      Result := ErrorResult(errOverflow);  // #NUM! as tested by Excel
      exit;
    end;
    base := ArgToFloat(Args[1]);
    if IsNaN(base) then begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    if base < 0 then begin
      Result := ErrorResult(errOverflow);  // #NUM!
      exit;
    end;
  end else
    base := 10;

  Result := FloatResult(logn(base, x));
end;

procedure fpsLOG10(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)    // #VALUE!
  else
  if x > 0 then
    Result := FloatResult(log10(x))
  else
    Result := ErrorResult(errOverflow);   // #NUM!
end;

procedure fpsMOD(var Result: TsExpressionResult; const Args: TsExprParameterArray);
//  MOD( number, divisor )
// Returns the remainder after a number is divided by a divisor.
var
  n, m: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  n := ArgToInt(Args[0]);
  m := ArgToInt(Args[1]);
  if m = 0 then
    Result := ErrorResult(errDivideByZero)
  else
    Result := IntegerResult(n mod m);
end;

procedure fpsODD(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// ODD( number )
// rounds a number up to the nearest odd integer.
// If the number is negative, the number is rounded away from zero.
var
  x: TsExprFloat;
  n: Integer;
begin
  if IsError(Args[0], Result) then
    exit;

  if Args[0].ResultType in [rtCell, rtInteger, rtFloat, rtDateTime, rtEmpty] then
  begin
    x := ArgToFloat(Args[0]);
    if IsNaN(x) then
      Result := ErrorResult(errWrongType)
    else
    if x >= 0 then
    begin
      n := Trunc(x) + 1;
      if not odd(n) then inc(n);
    end else
    begin
      n := Trunc(x) - 1;
      if not odd(n) then dec(n);
    end;
    Result := IntegerResult(n);
  end
  else
    Result := ErrorResult(errWrongType);
end;

procedure fpsPI(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Unused(Args);
  Result := FloatResult(pi);
end;

procedure fpsPOWER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x, y: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  y := ArgToFloat(Args[1]);
  if IsNaN(x) or IsNaN(y) then
    Result := ErrorResult(errWrongType)
  else
    try
      Result := FloatResult(Power(x, y));
    except
      Result := ErrorResult(errOverflow);
    end;
end;

procedure fpsRADIANS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(DegToRad(x));
end;

procedure fpsRAND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Unused(Args);
  Result := FloatResult(random);
end;

// Avoids Banker's rounding
function MyRoundTo(const AValue: Double; const Digits: TRoundToRange): Double;
var
  RV: Double;
begin
  RV := IntPower(10,Digits);
  Result := fpsUtils.Round(AValue / RV) * RV;
end;

procedure fpsROUND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
  n: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  x := ArgToFloat(Args[1]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else begin
    n := Round(x);
    x := ArgToFloat(Args[0]);
    if IsNaN(x) then
      Result := ErrorResult(errWrongType)
    else
      Result := FloatResult(MyRoundTo(x, -n));
      // -n because fpc and Excel have different conventions regarding the sign
  end;
end;

function MyRoundDown(const AValue: Double; const Digits: TRoundToRange): Double;
var
  RV, uValue, tmp, integral: Double;
begin
  uValue := abs(AValue);
  RV := IntPower(10, Digits);
  tmp := uValue * RV;
  integral := Int(tmp);
  if (1-Frac(tmp))/RV < EPS then
     tmp := integral +1
  else
     tmp := integral;
  Result := abs(tmp/RV)*sign(AValue);
end;

{ The Excel ROUNDDOWN function returns a number rounded down to a given number
  of decimal places. Unlike standard rounding, where only numbers less than 5
  are rounded down, ROUNDDOWN rounds all numbers down. }
procedure fpsROUNDDOWN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
  n: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  x := ArgToFloat(Args[1]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else begin
    n := Round(x);
    x := ArgToFloat(Args[0]);
    if IsNaN(x) then
      Result := ErrorResult(errWrongType)
    else
      Result := FloatResult(MyRoundDown(x, n));
  end;
end;

function MyRoundUp(const AValue: Double; const Digits: TRoundToRange): Double;
var
  RV, uValue, tmp, integral: Double;
begin
  uValue := abs(AValue);
  RV := IntPower(10, Digits);
  tmp := uValue * RV;
  integral := Int(tmp);
  if ((Frac(tmp)/RV) < EPS) then
    tmp := integral
  else
    tmp := integral +1;
  Result := abs(tmp/RV) * sign(AValue);
end;

{ The Excel ROUNDUP function returns a number rounded UP to a given number
  of decimal places. Unlike standard rounding, where only numbers less than 5
  are rounded UP, ROUNDUP rounds all numbers up. }
procedure fpsROUNDUP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
  n: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  x := ArgToFloat(Args[1]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else begin
    n := Round(x);
    x := ArgToFloat(Args[0]);
    if IsNaN(x) then
      Result := ErrorResult(errWrongType)
    else
      Result := FloatResult(MyRoundUp(x, n));
  end;
end;

procedure fpsSIGN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(sign(x));
end;

procedure fpsSIN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(sin(x));
end;

procedure fpsSINH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(sinh(x));
end;

procedure fpsSQRT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
  if x >= 0 then
    Result := FloatResult(sqrt(x))
  else
    Result := ErrorResult(errOverflow);
end;

procedure fpsTAN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
  if frac(x / (pi*0.5)) = 0 then
    Result := ErrorResult(errOverflow)   // #NUM!
  else
    Result := FloatResult(tan(x));
end;

procedure fpsTANH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  x: TsExprFloat;
begin
  if IsError(Args[0], Result) then
    exit;

  x := ArgToFloat(Args[0]);
  if IsNaN(x) then
    Result := ErrorResult(errWrongType)
  else
    Result := FloatResult(tanh(x));
end;


{------------------------------------------------------------------------------}
{   Built-in date/time functions                                               }
{------------------------------------------------------------------------------}

{ DATE( year, month, day )
  - If year is between 0 (zero) and 1899 (inclusive), Excel adds that value to
    1900 to calculate the year.
    Example, DATE(108,1,2) returns January 2, 2008 (1900+108).

  - If year is less than 0 or is 10000 or greater, Excel returns the #NUM! error value.

  - If month is greater than 12, month adds that number of months to the first
    month in the year specified.
    Example, DATE(2008,14,2) returns the serial number representing February 2, 2009.

  - If month is less than 1, month subtracts the magnitude of that number of
    months, plus 1, from the first month in the year specified.
    Example, DATE(2008,-3,2) returns the serial number representing September 2, 2007

  - day is greater than the number of days in the month specified, day adds
    that number of days to the first day in the month.
    Example, DATE(2008,1,35) returns the serial number representing February 4, 2008.

  - If day is less than 1, day subtracts the magnitude that number of days, plus one,
    from the first day of the month specified.
    Example, DATE(2008,1,-15) returns the serial number representing December 16, 2007
}
procedure fpsDATE(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
var
  refDate: TDate;
  yr, mn, dy, tmp, daysInMon: Integer;
  y, m, d: Word;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if IsError(Args[2], Result) then
    exit;

  yr := ArgToInt(Args[0]);
  mn := ArgToInt(Args[1]);
  dy := ArgToInt(Args[2]);

  // Handle 2-digit years and "overflow" years
  if (yr < 0) or (yr > 9999) then
  begin
    Result := ErrorResult(errOverFlow);
    exit;
  end;
  if (yr < 1900) then
    inc(yr, 1900);

  // Handle "normal, correct" date
  if InRange(mn, 1, 12) and InRange(dy, 1, DaysInMonth(mn)) then
  begin
    refDate := EncodeDate(yr, mn, dy);
  end
  else
  // No month overflow, but day overflow
  if InRange(mn, 1, 12) then
    refDate := EncodeDate(yr, mn, 1) + dy-1
  else
  begin
    // Month overflow
    y := yr;
    m := 1;
    d := 1;
    IncAMonth(y, m, d, mn-1);
    refDate := Encodedate(y, m, d);
    if InRange(dy, 1, DaysInMonth(m)) then  // No day overflow
      refDate := refdate + (dy - d)
    else // day overflow
      refDate := refdate + dy - 1;
  end;
  Result := DateTimeResult(refDate);
end;

{ DATEDIF( start_date, end_date, interval )
    start_date <= end_date !
    interval = Y  - The number of complete years.
             = M  - The number of complete months.
             = D  - The number of days.
             = MD - The difference between the days (months and years are ignored).
             = YM - The difference between the months (days and years are ignored).
             = YD - The difference between the days (years and dates are ignored). }
procedure fpsDATEDIF(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
var
  interval: String;
  start_date, end_date: TDate;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if IsError(Args[2], Result) then
    exit;

  start_date := ArgToDateTime(Args[0]);
  end_date := ArgToDateTime(Args[1]);
  if IsNaN(start_date) or IsNaN(end_date) then begin
    Result := ErrorResult(errWrongType);
    exit;
  end;
  interval := ArgToString(Args[2]);

  if end_date > start_date then
    Result := ErrorResult(errOverflow)
  else if interval = 'Y' then
    Result := FloatResult(YearsBetween(end_date, start_date))
  else if interval = 'M' then
    Result := FloatResult(MonthsBetween(end_date, start_date))
  else if interval = 'D' then
    Result := FloatResult(DaysBetween(end_date, start_date))
  else
    Result := ErrorResult(errFormulaNotSupported);
end;

// Returns the serial number of a date. Input is a string.
// DATE( date_string )
procedure fpsDATEVALUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  d: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  if TryStrToDate(Args[0].ResString, d) then
    Result := DateTimeResult(d)
  else
    Result := ErrorResult(errWrongType);      // #VALUE!
end;

// DAY( date_value )
// date_value can be a serial number or a string
procedure fpsDAY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  y,m,d: Word;
  dt: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger, rtString, rtCell]) then
  begin
    dt := ArgToDateTime(Args[0]);
    if not IsNaN(dt) then begin
      DecodeDate(dt, y, m, d);
      Result := IntegerResult(d);
    end;
  end;
end;

// HOUR( time_value )
// time_value can be a number or a string.
procedure fpsHOUR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  h, m, s, ms: Word;
  dt: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger, rtString, rtCell]) then
  begin
    dt := ArgToDateTime(Args[0]);
    if not IsNaN(dt) then begin
      DecodeTime(dt, h, m, s, ms);
      Result := IntegerResult(h);
    end;
  end;
end;

// MINUTE( serial_number or string )
procedure fpsMINUTE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  h, m, s, ms: Word;
  dt: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger, rtString, rtCell]) then
  begin
    dt := ArgToDateTime(Args[0]);
    if not IsNaN(dt) then begin
      DecodeTime(dt, h, m, s, ms);
      Result := IntegerResult(m);
    end;
  end;
end;

procedure fpsMONTH(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MONTH( date_value or string )
var
  y,m,d: Word;
  dt: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger, rtString, rtCell]) then
  begin
    dt := ArgToDateTime(Args[0]);
    if not IsNaN(dt) then begin
      DecodeDate(dt, y, m, d);
      Result := IntegerResult(m);
    end;
  end;
end;

procedure fpsNOW(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the current system date and time. Willrefresh the date/time value
// whenever the worksheet recalculates.
//   NOW()
begin
  Unused(Args);
  Result := DateTimeResult(Now);
end;

procedure fpsSECOND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// SECOND( serial_number )
var
  h, m, s, ms: Word;
  dt: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger, rtString, rtCell]) then
  begin
    dt := ArgToDateTime(Args[0]);
    if not IsNaN(dt) then begin
      DecodeTime(dt, h, m, s, ms);
      Result := IntegerResult(s);
    end;
  end;
end;

{ TIME( hour, minute, second)
  - hour: A number from 0 (zero) to 32767 representing the hour. Any value >23
     will be divided by 24 and the remainder will be treated as the hour value.
     For example, TIME(27,0,0) = TIME(3,0,0) = .125 or 3:00 AM.
  - minute: A number from 0 to 32767 representing the minute. Any value > 59
     will be converted to hours and minutes.
     For example, TIME(0,750,0) = TIME(12,30,0) = .520833 or 12:30 PM.
  - second: A number from 0 to 32767 representing the second. Any value > 59
     will be converted to hours, minutes, and seconds.
     For example, TIME(0,0,2000) = TIME(0,33,22) = .023148 or 12:33:20 AM }
procedure fpsTIME(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  hr, min, sec: Integer;
  t: Double;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if IsError(Args[2], Result) then
    exit;

  hr := ArgToInt(Args[0]);
  min := ArgToInt(Args[1]);
  sec := ArgToInt(Args[2]);

  t := (hr + min/60 + sec/3600) / 24;
  if t >= 0 then
    Result := DateTimeResult(t)
  else
    Result := ErrorResult(errOverFlow);  // #NUM!
end;

procedure fpsTIMEVALUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the serial number of a time. Input must be a string.
// DATE( date_string )
var
  t: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  if TryStrToTime(Args[0].ResString, t) then
    Result := DateTimeResult(t)
  else
    Result := ErrorResult(errWrongType);
end;

procedure fpsTODAY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the current system date. This function will refresh the date
// whenever the worksheet recalculates.
//   TODAY()
begin
  Unused(Args);
  Result := DateTimeResult(Date);
end;

procedure fpsWEEKDAY(var Result: TsExpressionResult; const Args: TsExprParameterArray);
{ WEEKDAY( serial_number, [return_value] )
   return_value = 1 - Returns a number from 1 (Sunday) to 7 (Saturday) (default)
                = 2 - Returns a number from 1 (Monday) to 7 (Sunday).
                = 3 - Returns a number from 0 (Monday) to 6 (Sunday). }
var
  n: Integer;
  dow: Integer;
  dt: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;
  Result := ErrorResult(errWrongType);
  if Length(Args) = 2 then
  begin
    if IsError(Args[1], Result) then
      exit;
    n := ArgToInt(Args[1]);
  end else
    n := 1;
  dt := NaN;
  if Args[0].ResultType in [rtDateTime, rtFloat, rtInteger, rtCell, rtString] then
    dt := ArgToDateTime(Args[0]);
  if IsNaN(dt) then
    exit;

  dow := DayOfWeek(dt);   // Sunday = 1 ... Saturday = 7
  case n of
    1: ;
    2: if dow > 1 then dow := dow - 1 else dow := 7;
    3: if dow > 1 then dow := dow - 2 else dow := 6;
  end;
  Result := IntegerResult(dow);
end;

procedure fpsYEAR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// YEAR( date_value )
var
  y,m,d: Word;
  dt: TDateTime;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  if (Args[0].ResultType in [rtDateTime, rtFloat, rtInteger, rtString, rtCell]) then
  begin
    dt := ArgToDateTime(Args[0]);
    if not IsNaN(dt) then begin
      DecodeDate(dt, y, m, d);
      Result := IntegerResult(y);
    end;
  end;
end;


{------------------------------------------------------------------------------}
{    Builtin string functions                                                  }
{------------------------------------------------------------------------------}

procedure fpsCHAR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CHAR( ascii_value )
// returns the character based on the ASCII value
var
  arg: Integer;
begin
  if IsError(Args[0], Result) then
    exit;

  Result := ErrorResult(errWrongType);
  case Args[0].ResultType of
    rtInteger, rtFloat:
      if Args[0].ResultType in [rtInteger, rtFloat] then
      begin
        arg := ArgToInt(Args[0]);
        if (arg >= 0) and (arg < 256) then
          Result := StringResult(AnsiToUTF8(Char(arg)));
      end;
    rtError:
      Result := ErrorResult(Args[0].ResError);
    rtEmpty:
      Result.ResultType := rtEmpty;
  end;
end;

procedure fpsCODE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CODE( text )
// returns the ASCII value of a character or the first character in a string.
var
  s: String;
  ch: Char;
begin
  if IsError(Args[0], Result) then
    exit;

  s := ArgToString(Args[0]);
  if s = '' then
    Result := ErrorResult(errWrongType)
  else
  begin
    ch := UTF8ToAnsi(s)[1];
    Result := IntegerResult(ord(ch));
  end;
end;

procedure fpsCONCATENATE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CONCATENATE( text1, text2, ... text_n )
// Joins two or more strings together
var
  s: String;
  i: Integer;
begin
  s := '';
  for i:=0 to Length(Args)-1 do
  begin
    if IsError(Args[i], Result) then
      exit;
    s := s + ArgToString(Args[i]);
  end;
  Result := StringResult(s);
end;

procedure fpsEXACT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// EXACT( text1, text2 )
// Compares two strings (case-sensitive) and returns TRUE if they are equal
var
  s1, s2: String;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;

  s1 := ArgToString(Args[0]);
  s2 := ArgToString(Args[1]);
  Result := BooleanResult(s1 = s2);
end;

procedure fpsLEFT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LEFT( text, [number_of_characters] )
// extracts a substring from a string, starting from the left-most character
var
  s: String;
  count: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if (Length(Args) > 1) and IsError(Args[1], Result) then
    exit;

  s := ArgToString(Args[0]);
  if s = '' then
    Result := EmptyResult
  else
  if Length(Args) > 2 then
    Result := ErrorResult(errArgError)
  else
  begin
    if Length(Args) = 1 then
      count := 1
    else
    if Args[1].ResultType = rtMissingArg then
      count := 1
    else
      count := ArgToInt(Args[1]);
    Result := StringResult(UTF8LeftStr(s, count));
  end;
end;

procedure fpsLEN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LEN( text )
//  returns the length of the specified string.
begin
  if IsError(Args[0], Result) then
    exit;
  Result := IntegerResult(UTF8Length(ArgToString(Args[0])));
end;

procedure fpsLOWER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// LOWER( text )
// converts all letters in the specified string to lowercase. If there are
// characters in the string that are not letters, they are not affected.
begin
  if IsError(Args[0], Result) then
    exit;
  Result := StringResult(UTF8Lowercase(ArgToString(Args[0])));
end;

procedure fpsMID(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MID( text, start_position, number_of_characters )
// extracts a substring from a string (starting at any position).
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if IsError(Args[2], Result) then
    exit;
  Result := StringResult(UTF8Copy(ArgToString(Args[0]), ArgToInt(Args[1]), ArgToInt(Args[2])));
end;

procedure fpsREPLACE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// REPLACE( old_text, start, number_of_chars, new_text )
//  replaces a sequence of characters in a string with another set of characters
var
  sOld, sNew, s1, s2: String;
  start: Integer;
  count: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if IsError(Args[2], Result) then
    exit;
  if IsError(Args[3], Result) then
    exit;
  sOld := Args[0].ResString;
  start := ArgToInt(Args[1]);
  count := ArgToInt(Args[2]);
  sNew := Args[3].ResString;
  s1 := UTF8Copy(sOld, 1, start-1);
  s2 := UTF8Copy(sOld, start+count, UTF8Length(sOld));
  Result := StringResult(s1 + sNew + s2);
end;

procedure fpsREPT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// REPT( text, count )
// repeats text a specified number of times.
var
  s: String;
  count: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  s := ArgToString(Args[0]);
  if s = '' then
    Result.ResultType := rtEmpty
  else
  if Args[1].ResultType in [rtInteger, rtFloat] then begin
    count := ArgToInt(Args[1]);
    Result := StringResult(DupeString(s, count));
  end;
end;

procedure fpsRIGHT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// RIGHT( text, [number_of_characters] )
// extracts a substring from a string, starting from the last character
var
  s: String;
  count: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if (Length(Args) > 1) and IsError(Args[1], Result) then
    exit;

  s := ArgToString(Args[0]);
  if s = '' then
    Result := EmptyResult
  else
  if Length(Args) > 2 then
    Result := ErrorResult(errArgError)
  else
  begin
    if Length(Args) = 1 then
      count := 1
    else
    if Args[1].ResultType = rtMissingArg then
      count := 1
    else
      count := ArgToInt(Args[1]);
    Result := StringResult(UTF8RightStr(s, count));
  end;
end;

procedure fpsSUBSTITUTE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// SUBSTITUTE( text, old_text, new_text, [nth_appearance] )
// replaces a set of characters with another.
var
  sOld: String;
  sNew: String;
  s1, s2: String;
  n: Integer;
  s: String;
  p: Integer;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if IsError(Args[2], Result) then
    exit;
  if (Length(Args) > 3) and IsError(Args[3], Result) then
    exit;
  sOld := ArgToString(Args[1]);
  sNew := ArgToString(Args[2]);
  if Length(Args) = 4 then
  begin
    if IsError(Args[3], Result) then
      exit;
    n := ArgToInt(Args[3]);               // THIS PART NOT YET CHECKED !!!!!!
    if n <= 0 then
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    p := UTF8Pos(sOld, s);
    while (n > 1) do begin
      p := UTF8Pos(sOld, s, p+1);
      dec(n);
    end;
    if p > 0 then begin
      s1 := UTF8Copy(s, 1, p-1);
      s2 := UTF8Copy(s, p+UTF8Length(sOld), UTF8Length(s));
      s := s1 + sNew + s2;
    end;
    Result := StringResult(s);
  end else
    Result := StringResult(UTF8StringReplace(s, sOld, sNew, [rfReplaceAll]));
end;

procedure fpsTEXT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// TEXT( value, format )
// Returns a value converted to text with a specified format.
var
  fmt: String;
  value: double;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  value := ArgToFloat(Args[0]);
  fmt := ArgToString(Args[1]);
  if IsDateTimeFormat(fmt) then
    Result := StringResult(FormatDateTime(fmt, value))
  else
    Result := StringResult(Format(fmt, [value]));
end;

procedure fpsTRIM(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// TRIM( text )
// Returns a text value with the leading and trailing spaces removed
begin
  if IsError(Args[0], Result) then
    exit;
  Result := StringResult(UTF8Trim(ArgToString(Args[0])));
end;

procedure fpsUPPER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// UPPER( text )
// converts all letters in the specified string to uppercase. If there are
// characters in the string that are not letters, they are not affected.
begin
  if IsError(Args[0], Result) then
    exit;
  Result := StringResult(UTF8Uppercase(ArgToString(Args[0])));
end;

procedure fpsVALUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// VALUE( text )
// converts a text value that represents a number to a number.
var
  x: Double;
  n: Integer;
  s: String;
begin
  if IsError(Args[0], Result) then
    exit;
  s := ArgToString(Args[0]);
  if TryStrToInt(s, n) then
    Result := IntegerResult(n)
  else
  if TryStrToFloat(s, x, ExprFormatSettings) then
    Result := FloatResult(x)
  else
  if TryStrToDateTime(s, x) then
    Result := FloatResult(x)
  else
    Result := ErrorResult(errWrongType);
end;


{------------------------------------------------------------------------------}
{    Built-in logical functions                                                }
{------------------------------------------------------------------------------}

// AND( condition1, [condition2], ... )
// up to 30 parameters. At least 1 parameter.
procedure fpsAND(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  i: Integer;
  b: Boolean;
begin
  b := true;
  for i:=0 to High(Args) do
  begin
    if IsError(Args[i], Result) then
      exit;
    if (Args[0].ResultType = rtString) then
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    if not ArgToBoolean(Args[i], false) then begin
      b := false;
      break;
    end;
  end;
  Result.ResBoolean := b;
end;

// FALSE ()
procedure fpsFALSE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Unused(Args);
  Result.ResBoolean := false;
end;

// IF( condition, value_if_true, [value_if_false] )
procedure fpsIF(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  res: Boolean;
begin
  if IsError(Args[0], Result) then
    exit;

  if (Args[0].ResultType = rtString) then
  begin
    Result := ErrorResult(errWrongType);
    exit;
  end;

  res := ArgToBoolean(Args[0], false);
  if res then
    Result := Args[1]
  else
  begin
    if Length(Args) > 2 then
      Result := Args[2]
    else
      Result.ResBoolean := false;
  end;
end;

// IFS( condition, value_if_true, [condition], [value_if_true], [condition], [value_if_true] )
procedure fpsIFS(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  i:integer;
begin
  Result := ErrorResult(errArgError);
  if (Length(Args) mod 2 <> 0) then   // We always need pairs of args
    exit;                             // --> If not, exit with argument eror
  i:=0;
  while(i < Length(Args)-1) do begin
    if IsError(Args[i], Result) then
      exit;
    if (Args[0].ResultType = rtString) then  // A string never represents a boolean value in Excel
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    if ArgToBoolean(Args[i], false) then
    begin
      Result := Args[i+1];
      break;
    end; // What if ArgToBoolean is false?
    inc(i, 2);
  end;
end;

// NOT( condition )
procedure fpsNOT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  if IsError(Args[0], Result) then
    exit;
  if (Args[0].ResultType = rtString) then
    Result := ErrorResult(errWrongType)
  else
    Result.ResBoolean := not ArgToBoolean(Args[0], false);
end;

// OR( condition1, [condition2], ... )
// up to 30 parameters. At least 1 parameter.
procedure fpsOR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  i: Integer;
  b: Boolean;
begin
  b := false;
  for i:=0 to High(Args) do
  begin
    if IsError(Args[i], Result) then
      exit;
    if (Args[0].ResultType = rtString) then
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;
    if ArgToBoolean(Args[i], false) then begin
      b := true;
      break;
    end;
  end;
  Result.ResBoolean := b;
end;

// TRUE()
procedure fpsTRUE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  Unused(Args);
  Result.ResBoolean := true;
end;


{------------------------------------------------------------------------------}
{    Built-in statistical functions                                            }
{------------------------------------------------------------------------------}

// Average value of absolute deviations of data from their mean.
// AVEDEV( value1, [value2, ... value_n] )
procedure fpsAVEDEV(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  data: TsExprFloatArray = nil;
  m: TsExprFloat;
  i: Integer;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  for i := 0 to Length(Args)-1 do
    if IsError(Args[i], Result) then
      exit;

  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);

  if err = errOK then
  begin
    if Length(data) = 0 then
    begin
      Result := ErrorResult(errWrongType);
      exit;
    end;

    m := Mean(data);
    for i:=0 to High(data) do      // replace data by their average deviation from the mean
      data[i] := abs(data[i] - m);
    Result.ResFloat := Mean(data);
  end else
    Result := ErrorResult(err);
end;

procedure fpsAVERAGE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// AVERAGE( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);

  if (Length(data) = 0) then
    if not hasLiteralStrings then
    begin
      Result := ErrorResult(errDivideByZero);
      exit;
    end;

  if err = errOK then
    Result.ResFloat := Mean(data)
  else
    Result := ErrorResult(err);
end;

{ counts the number of cells that contain numbers as well as the number of
  arguments that contain numbers.
    COUNT( value1, [value2, ... value_n] )  }
procedure fpsCOUNT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, false, data, err, hasLiteralStrings);
  Result := IntegerResult(Length(data));
end;

// Counts the number of cells that are not empty as well as the number of
// arguments that contain values
//   COUNTA( value1, [value2, ... value_n] )
procedure fpsCOUNTA(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  i, n: Integer;
  r, c: Cardinal;
  cell: PCell;
  arg: TsExpressionResult;
  book: TsWorkbook;
  sheet: TsWorksheet;
  sheetIdx: Integer;
begin
  n := 0;
  for i:=0 to High(Args) do
  begin
    arg := Args[i];
    case arg.ResultType of
      rtInteger, rtFloat, rtDateTime, rtBoolean:
        inc(n);
      rtString:
        if arg.ResString <> '' then inc(n);
      rtError:
        if arg.ResError <> errOK then inc(n);
      rtCell:
        begin
          cell := ArgToCell(arg);
          if cell <> nil then
            case cell^.ContentType of
              cctNumber, cctDateTime, cctBool: inc(n);
              cctUTF8String: if cell^.UTF8StringValue <> '' then inc(n);
              cctError: if cell^.ErrorValue <> errOK then inc(n);
            end;
        end;
      rtCellRange:
        begin
          book := TsWorkbook(TsWorksheet(arg.Worksheet).Workbook);
          for sheetIdx := arg.ResCellRange.Sheet1 to arg.ResCellRange.Sheet2 do
          begin
            sheet := book.GetWorksheetByIndex(sheetIdx);
            for r := arg.ResCellRange.Row1 to arg.ResCellRange.Row2 do
              for c := arg.ResCellRange.Col1 to arg.ResCellRange.Col2 do
              begin
                cell := sheet.FindCell(r, c);
                if (cell <> nil) then
                  case cell^.ContentType of
                    cctNumber, cctDateTime, cctBool : inc(n);
                    cctUTF8String: if cell^.UTF8StringValue <> '' then inc(n);
                    cctError: if cell^.ErrorValue <> errOK then inc(n);
                  end;
              end;
          end;
        end;
    end;
  end;
  Result.ResInteger := n;
end;

{ Counts the number of empty cells in a range.
    COUNTBLANK( range )
  "range" is the range of cells to count empty cells. }
procedure fpsCOUNTBLANK(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  n: Integer;
  r, c: Cardinal;
  cell: PCell;
begin
  if IsError(Args[0], Result) then
    exit;
  n := 0;
  case Args[0].ResultType of
    rtEmpty:
      inc(n);
    rtCell:
      begin
        cell := ArgToCell(Args[0]);
        if cell = nil then
          inc(n)
        else
          case cell^.ContentType of
            cctNumber, cctDateTime, cctBool: ;
            cctUTF8String: if cell^.UTF8StringValue = '' then inc(n);
            cctError: if cell^.ErrorValue = errOK then inc(n);
          end;
      end;
    rtCellRange:
      for r := Args[0].ResCellRange.Row1 to Args[0].ResCellRange.Row2 do
        for c := Args[0].ResCellRange.Col1 to Args[0].ResCellRange.Col2 do begin
          cell := (Args[0].Worksheet as TsWorksheet).FindCell(r, c);
          if cell = nil then
            inc(n)
          else
            case cell^.ContentType of
              cctNumber, cctDateTime, cctBool: ;
              cctUTF8String: if cell^.UTF8StringValue = '' then inc(n);
              cctError: if cell^.ErrorValue = errOK then inc(n);
            end;
        end;
  end;
  Result.ResInteger := n;
end;

{ Calculates the average value of the cell values if they meet a given condition.
    AVERAGEIF( range, condition, [ave_range] )
  - "range" is the cell range to be analyzed
  - "condition" can be a cell, a value or a string starting with a symbol like ">" etc.
    (in the former two cases a value is counted if equal to the criteria value)
  - "ave_range" - option for the values to be added; if missing the values in
    "range" are used.}
procedure fpsAVERAGEIF(var AResult: TsExpressionresult; const Args: TsExprParameterArray);
var
  cmp: TsFuncComparer;
  vr, cr, c: Integer;
begin
  cr := 0;   // (First) criteria range arg index
  c := 1;    // (First) criteria arg index.
  if Length(Args) = 2 then vr := 0 else vr := 2;  // value range arg index
  cmp := TsFuncComparer.Create(Args, vr, cr, c, ftAVERAGEIF);
  try
    AResult := cmp.Execute;
  finally
    cmp.Free;
  end;
end;

procedure fpsAVERAGEIFS(var AResult: TsExpressionresult; const Args: TsExprParameterArray);
{ Calculates the average value of the cell values if they meet the specified condition.
    AVERAGEIFS(avg_range, criteria_range, criteria [, criteria_range, criteria, ...])
  - "avg_range" is the range in which the average value is calculated
  - "criteria_range" is the cell range to be compared with the criteria
  - "criteria" can be a cell, a value or a string starting with a symbol like ">" etc.
    (in the former two cases a value is counted if equal to the criteria value)
}
var
  cmp: TsFuncComparer;
begin
  cmp := TsFuncComparer.Create(Args, 0, 1, 2, ftAVERAGEIFS);
  try
    AResult := cmp.Execute;
  finally
    cmp.Free;
  end;
end;

{ Counts the number of cells in a range that meets a given condition.
    COUNTIF( range, condition )
  - "range" is the cell range to be analyzed
  - "condition" can be a cell, a value or a string starting with a symbol like ">" etc.
    (in the former two cases a value is counted if equal to the criteria value) }
procedure fpsCOUNTIF(var AResult: TsExpressionResult; const Args: TsExprParameterArray);
var
  cmp: TsFuncComparer;
begin
  cmp := TsFuncComparer.Create(Args, 0, 0, 1, ftCOUNTIF);
  try
    AResult := cmp.Execute;
  finally
    cmp.Free;
  end;
end;

{ Calculates the count value of the (numeric) cell values if they meet the specified condition.
    COUNTIFS(count_range, criteria_range, criteria [, criteria_range, criteria, ...])
  - "count_range" is the range over which the cells are counted.
  - "criteria_range" is the cell range to be compared with the criteria
  - "criteria" can be a cell, a value or a string starting with a symbol like ">" etc.
    (in the former two cases a value is counted if equal to the criteria value)
}
procedure fpsCOUNTIFS(var AResult: TsExpressionresult; const Args: TsExprParameterArray);
var
  cmp: TsFuncComparer;
begin
  cmp := TsFuncComparer.Create(Args, 0, 0, 1, ftCOUNTIFS);
  try
    AResult := cmp.Execute;
  finally
    cmp.Free;
  end;
end;

{ Adds the cell values if they meet a given condition.
    SUMIF( range, condition, [sum_range] )
  - "range" is the cell range to be analyzed
  - "condition" can be a cell, a value or a string starting with a symbol like ">" etc.
    (in the former two cases a value is added if equal to the criteria value)
  - "sum_range" - option for the values to be added; if missing the values in
    "range" are used.}
procedure fpsSUMIF(var AResult: TsExpressionResult; const Args: TsExprParameterArray);
var
  cmp: TsFuncComparer;
  vr, cr, c: Integer;
begin
  if IsError(Args[0], AResult) then
    exit;
  { Excel still calculates the formula if the "condition" argument contains an error
  if IsError(Args[1], AResult) then
    exit;
  }
  if (Length(Args) = 3) and IsError(Args[0], AResult) then
    exit;

  cr := 0;   // First criteria range arg index
  c := 1;    // First criteria arg index
  if Length(Args) = 2 then
    vr := 0  // value range index
  else
    vr := 2;
  cmp := TsFuncComparer.Create(Args, vr, cr, c, ftSUMIF);
  try
    AResult := cmp.Execute;
  finally
    cmp.Free;
  end;
end;

procedure fpsSUMIFS(var AResult: TsExpressionresult; const Args: TsExprParameterArray);
{ Calculates the sum of the cell values if they meet the specified condition.
    SUMIFS(sum_range, criteria_range, criteria [, criteria_range, criteria, ...])
  - "sum_range" is the range over which the cells are summed.
  - "criteria_range" is the cell range to be compared with the criteria
  - "criteria" can be a cell, a value or a string starting with a symbol like ">" etc.
    (in the former two cases a value is counted if equal to the criteria value)
}
var
  cmp: TsFuncComparer;
begin
  cmp := TsFuncComparer.Create(Args, 0, 1, 2, ftSUMIFS);
  try
    AResult := cmp.Execute;
  finally
    cmp.Free;
  end;
end;

procedure fpsMAX(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MAX( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err = errOK then
  begin
    if Length(data) > 0 then
      Result.ResFloat := MaxValue(data)
    else
      Result.ResFloat := 0;
  end else
    Result := ErrorResult(err);
end;

procedure fpsMIN(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// MIN( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err = errOK then
  begin
    if Length(data) > 0 then
      Result.ResFloat := MinValue(data)
    else
      Result.ResFloat := 0;
  end else
    Result := ErrorResult(err);
end;

procedure fpsPRODUCT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// PRODUCT( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray = nil;
  i: Integer;
  p: TsExprFloat;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err <> errOK then begin
    Result := ErrorResult(err);
    exit;
  end;

  if Length(data) = 0 then
    Result.ResFloat := 0.0
  else
  begin
    p := 1.0;
    for i := 0 to High(data) do
      p := p * data[i];
    Result.ResFloat := p;
  end;
end;

procedure fpsSTDEV(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the standard deviation of a population based on a sample of numbers
// of numbers.
//   STDEV( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err <> errOK then begin
    Result := ErrorResult(err);
    exit;
  end;

  if Length(data) > 1 then
    Result.ResFloat := StdDev(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;

procedure fpsSTDEVP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the standard deviation of a population based on an entire population
// STDEVP( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err <> errOK then begin
    Result := ErrorResult(err);
    exit;
  end;

  if Length(data) > 0 then
    Result.ResFloat := PopnStdDev(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;

// SUM( value1, [value2, ... value_n] )
procedure fpsSUM(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err = errOK then
    Result.ResFloat := Sum(data)
  else
    Result := ErrorResult(err);
end;

// Returns the sum of the squares of a series of values.
// SUMSQ( value1, [value2, ... value_n] )
procedure fpsSUMSQ(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err = errOK then
    Result.ResFloat := SumOfSquares(data)
  else
    Result := ErrorResult(err);
end;

// Returns the variance of a population based on a sample of numbers.
// VAR( value1, [value2, ... value_n] )
procedure fpsVAR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err <> errOK then
  begin
    Result := ErrorResult(err);
    exit;
  end;

  if Length(data) > 1 then
    Result.ResFloat := Variance(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;

procedure fpsVARP(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// Returns the variance of a population based on an entire population of numbers.
// VARP( value1, [value2, ... value_n] )
var
  data: TsExprFloatArray = nil;
  err: TsErrorValue;
  hasLiteralStrings: Boolean;
begin
  ArgsToFloatArray(Args, true, data, err, hasLiteralStrings);
  if err <> errOK then
  begin
    Result := ErrorResult(err);
    exit;
  end;

  if Length(data) > 0 then
    Result.ResFloat := PopnVariance(data)
  else
  begin
    Result.ResultType := rtError;
    Result.ResError := errDivideByZero;
  end;
end;


{------------------------------------------------------------------------------}
{    Builtin info functions                                                    }
{------------------------------------------------------------------------------}

{ !!!!!!!!!!!!!!  not working !!!!!!!!!!!!!!!!!!!!!! }
{ !!!!!!!!!!!!!! needs localized strings !!!!!!!!!!! }

procedure fpsCELL(var Result: TsExpressionResult; const Args: TsExprParameterArray);
// CELL( type, [range] )

{ from http://www.techonthenet.com/excel/formulas/cell.php:

  "type" is the type of information that we retrieve for the cell and can have
  one of the following values:
    Value         Explanation
    ------------- --------------------------------------------------------------
    "address"     Address of the cell. If the cell refers to a range, it is the
                  first cell in the range.
    "col"         Column number of the cell.
    "color"       Returns 1 if the color is a negative value; Otherwise it returns 0.
    "contents"    Contents of the upper-left cell.
    "filename"    Filename of the file that contains reference.
    "format"      Number format of the cell according to:
                     "G"    General
                     "F0"   0
                     ",0"   #,##0
                     "F2"   0.00
                     ",2"   #,##0.00
                     "C0"   $#,##0_);($#,##0)
                     "C0-"  $#,##0_);[Red]($#,##0)
                     "C2"   $#,##0.00_);($#,##0.00)
                     "C2-"  $#,##0.00_);[Red]($#,##0.00)
                     "P0"   0%
                     "P2"   0.00%
                     "S2"   0.00E+00
                     "G"    # ?/? or # ??/??
                     "D4"   m/d/yy or m/d/yy h:mm or mm/dd/yy
                     "D1"   d-mmm-yy or dd-mmm-yy
                     "D2"   d-mmm or dd-mmm
                     "D3"   mmm-yy
                     "D5"   mm/dd
                     "D6"   h:mm:ss AM/PM
                     "D7"   h:mm AM/PM
                     "D8"   h:mm:ss
                     "D9"   h:mm
    "parentheses" Returns 1 if the cell is formatted with parentheses;
                  Otherwise, it returns 0.
    "prefix"      Label prefix for the cell.
                  - Returns a single quote (') if the cell is left-aligned.
                  - Returns a double quote (") if the cell is right-aligned.
                  - Returns a caret (^) if the cell is center-aligned.
                  - Returns a back slash (\) if the cell is fill-aligned.
                  - Returns an empty text value for all others.
    "protect"     Returns 1 if the cell is locked. Returns 0 if the cell is not locked.
    "row"         Row number of the cell.
    "type"        Returns "b" if the cell is empty.
                  Returns "l" if the cell contains a text constant.
                  Returns "v" for all others.
    "width"       Column width of the cell, rounded to the nearest integer.

  !!!! NOT ALL OF THEM ARE SUPPORTED HERE !!!

  "range" is optional in Excel. It is the cell (or range) that you wish to retrieve
  information for. If the range parameter is omitted, the CELL function will
  assume that you are retrieving information for the last cell that was changed.

  "range" is NOT OPTIONAL here because we don't know the last cell changed !!!
}
var
  stype: String;
  r1, c1: Cardinal;
  cell: PCell;
  cellfmt: TsCellFormat;
begin
  if IsError(Args[0], Result) then
    exit;
  if (Length(Args) = 2) and IsError(Args[1], Result) then
    exit;

  if Length(Args)=1 then
  begin
    // This case is not supported by us, but it is by Excel.
    // Therefore the error is not quite correct...
    Result := ErrorResult(errIllegalRef);
    exit;
  end;

  stype := lowercase(ArgToString(Args[0]));

  case Args[1].ResultType of
    rtCell:
      begin
        cell := ArgToCell(Args[1]);
        r1 := Args[1].ResRow;
        c1 := Args[1].ResCol;
      end;
    rtCellRange:
      begin
        r1 := Args[1].ResCellRange.Row1;
        c1 := Args[1].ResCellRange.Col1;
        cell := (Args[1].Worksheet as TsWorksheet).FindCell(r1, c1);
      end;
    else
      Result := ErrorResult(errWrongType);
      exit;
  end;
  if cell <> nil then
    cellfmt := TsWorksheet(cell^.Worksheet).ReadCellFormat(cell)
  else
    InitFormatRecord(cellfmt);

  if stype = 'address' then
    Result := StringResult(GetCellString(r1, c1, []))
  else
  if stype = 'col' then
    Result := IntegerResult(c1+1)
  else
  if stype = 'color' then
  begin
    if (cell <> nil) and (cellfmt.NumberFormat = nfCurrencyRed) then
      Result := IntegerResult(1)
    else
      Result := IntegerResult(0);
  end else
  if stype = 'contents' then
  begin
    if cell = nil then
      Result := IntegerResult(0)
    else
      case cell^.ContentType of
        cctNumber     : if frac(cell^.NumberValue) = 0 then
                          Result := IntegerResult(trunc(cell^.NumberValue))
                        else
                          Result := FloatResult(cell^.NumberValue);
        cctDateTime   : Result := DateTimeResult(cell^.DateTimeValue);
        cctUTF8String : Result := StringResult(cell^.UTF8StringValue);
        cctBool       : Result := BooleanResult(cell^.BoolValue);
        cctError      : Result := ErrorResult(cell^.ErrorValue);
      end;
  end else
  if stype = 'filename' then
    Result := Stringresult(
      ExtractFilePath((Args[1].Worksheet as TsWorksheet).Workbook.FileName) + '[' +
      ExtractFileName((Args[1].Worksheet as TsWorksheet).Workbook.FileName) + ']' +
      Args[1].Worksheet.Name
    )
  else
  if stype = 'format' then begin
    Result := StringResult('G');
    if cell <> nil then
      case cellfmt.NumberFormat of
        nfGeneral:
          Result := StringResult('G');
        nfFixed:
          if cellfmt.NumberFormatStr= '0' then Result := StringResult('0') else
          if cellfmt.NumberFormatStr = '0.00' then  Result := StringResult('F0');
        nfFixedTh:
          if cellfmt.NumberFormatStr = '#,##0' then Result := StringResult(',0') else
          if cellfmt.NumberFormatStr = '#,##0.00' then Result := StringResult(',2');
        nfPercentage:
          if cellfmt.NumberFormatStr = '0%' then Result := StringResult('P0') else
          if cellfmt.NumberFormatStr = '0.00%' then Result := StringResult('P2');
        nfExp:
          if cellfmt.NumberFormatStr = '0.00E+00' then Result := StringResult('S2');
        nfShortDate, nfLongDate, nfShortDateTime:
          Result := StringResult('D4');
        nfLongTimeAM:
          Result := StringResult('D6');
        nfShortTimeAM:
          Result := StringResult('D7');
        nfLongTime:
          Result := StringResult('D8');
        nfShortTime:
          Result := StringResult('D9');
      end;
  end else
  if stype = 'prefix' then
  begin
    Result := StringResult('');
    if (cell^.ContentType = cctUTF8String) then
      case cellfmt.HorAlignment of
        haLeft  : Result := StringResult('''');
        haCenter: Result := StringResult('^');
        haRight : Result := StringResult('"');
      end;
  end else
  if stype = 'row' then
    Result := IntegerResult(r1+1)
  else
  if stype = 'type' then begin
    if (cell = nil) or (cell^.ContentType = cctEmpty) then
      Result := StringResult('b')
    else if cell^.ContentType = cctUTF8String then begin
      if (cell^.UTF8StringValue = '')
        then Result := StringResult('b')
        else Result := StringResult('l');
    end else
      Result := StringResult('v');
  end else
  if stype = 'width' then
    Result := FloatResult((Args[1].Worksheet as TsWorksheet).GetColWidth(c1, suChars))
  else
    Result := ErrorResult(errWrongType);
end;

//  ERROR.TYPE(value)
// returns the numeric representation of one of the errors in Excel.
// "value" can be one of the following Excel error values
//   #NULL!  #DIV/0!  #VALUE!  #REF!  #NAME?  #NUM!  #N/A  // #GETTING_DATA and newer ones -- not supported
// When there is no error in the argument the function reports a #N/A error!
procedure fpsERRORTYPE(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  if (Args[0].ResultType = rtError) and (ord(Args[0].ResError) <= ord(errArgError))
  then
    Result := IntegerResult(ord(Args[0].ResError))
  else
    Result := ErrorResult(errArgError);
end;

//  ISBLANK( value )
// Checks for blank or null values.
// "value" is the value that you want to test.
// If "value" is blank, this function will return TRUE.
// If "value" is not blank, the function will return FALSE.
procedure fpsISBLANK(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  // No check for errors in Args here!

  Result := BooleanResult(false);
  case Args[0].ResultType of
    rtEmpty : Result := BooleanResult(true);
//    rtString: Result := BooleanResult(Args[0].ResString = '');  --> Excel returns false here!
    rtCell  : begin
                cell := ArgToCell(Args[0]);
                if (cell = nil) or (cell^.ContentType = cctEmpty) then
                  Result := BooleanResult(true);
              end;
  end;
end;

// ISERR( value )
// If "value" is an error value (except #N/A), this function will return TRUE.
// Otherwise, it will return FALSE.
procedure fpsISERR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctError) and (cell^.ErrorValue <> errArgError)
      then Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtError) and (Args[0].ResError <> errArgError) then
    Result := BooleanResult(true);
end;

// ISERROR( value )
// If "value" is an error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?
// or #NULL), this function will return TRUE. Otherwise, it will return FALSE.
procedure fpsISERROR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctError) and (cell^.ErrorValue <= errArgError)
      then Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtError) then
    Result := BooleanResult(true);
end;

// IFERROR( value; value_if_error )
// If "value" is an error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?
// or #NULL), this function will return Args[1]. Otherwise, it will return Args[0].
procedure fpsIFERROR(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctError) and (cell^.ErrorValue <= errArgError)
      then Result := Args[1]
    else result := Args[0];
  end else
  if (Args[0].ResultType = rtError) then
    Result := Args[1]
  else
    Result := Args[0];
end;

// ISLOGICAL( value )
procedure fpsISLOGICAL(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  // No check for errors in Args here!

  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctBool) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtBoolean) then
    Result := BooleanResult(true);
end;

// ISNA( value )
// If "value" is a #N/A error value , this function will return TRUE.
// Otherwise, it will return FALSE.
procedure fpsISNA(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  // No check for errors in Args here!

  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctError) and (cell^.ErrorValue = errArgError)
      then Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtError) and (Args[0].ResError = errArgError) then
    Result := BooleanResult(true);
end;

// ISNONTEXT( value )
procedure fpsISNONTEXT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  // No check for errors in Args here!

  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell = nil) or ((cell <> nil) and (cell^.ContentType <> cctUTF8String)) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType <> rtString) then
    Result := BooleanResult(true);
end;

// ISNUMBER( value )
// Tests "value" for a number (or date/time - checked with Excel).
procedure fpsISNUMBER(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  // No check for errors in Args here!

  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType in [cctNumber, cctDateTime]) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType in [rtFloat, rtInteger, rtDateTime]) then
    Result := BooleanResult(true);
end;

// ISREF( value )
procedure fpsISREF(var Result: TsExpressionResult; const Args: TsExprParameterArray);
begin
  // No check for errors in Args here!
  Result := BooleanResult(Args[0].ResultType in [rtCell, rtCellRange]);
end;

// ISTEXT( value )
procedure fpsISTEXT(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  cell: PCell;
begin
  // No check for errors in Args here!

  Result := BooleanResult(false);
  if (Args[0].ResultType = rtCell) then
  begin
    cell := ArgToCell(Args[0]);
    if (cell <> nil) and (cell^.ContentType = cctUTF8String) then
      Result := BooleanResult(true);
  end else
  if (Args[0].ResultType = rtString) then
    Result := BooleanResult(true);
end;


{------------------------------------------------------------------------------}
{    Builtin lookup/reference functions                                        }
{------------------------------------------------------------------------------}

{ ADDRESS(row, column, [ref_type], [ref_style], [sheet_name] )
  Returns a text representation of a cell address.
  "row" and "column": row and column indices, 1-based.
  "ref_type" is the type of reference to use: 1=absolute, 2=rel col, abs row,
    3= abs col, rel row, 4=relative; if omitted, 1 (absolute) is assumed.
  "ref_style" if true (default) means: address in A1 dialect, otherwise in R1C1.
  "sheet_name": name of the worksheet. Note, when sheet_name is used the
    address is presented in Excel dialects only. }
procedure fpsADDRESS(var Result: TsExpressionResult;
  const ARgs: TsExprParameterArray);
var
  c, r: Integer;
  flags: TsRelFlags;
  sheet: String;
  resStr: String;
  A1Dialect: Boolean;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if (Length(Args) > 2) and IsError(Args[2], Result) then
    exit;
  if (Length(Args) > 3) and IsError(Args[3], Result) then
    exit;
  if (Length(Args) > 4) and IsError(Args[4], Result) then
    exit;

  Result := ErrorResult(errArgError);
  if Length(Args) < 2 then
    exit;
  r := ArgToInt(Args[0]) - 1;
  c := ArgToInt(Args[1]) - 1;

  flags := [];
  if Length(Args) > 2 then
    case ArgToInt(Args[2]) of
      1: ;
      2: flags := [rfRelCol];
      3: flags := [rfRelRow];
      4: flags := [rfRelCol, rfRelRow];
    end;

  A1Dialect := true;
  if (Length(Args) > 3) and (Args[3].ResultType <> rtMissingArg) then
    A1Dialect := ArgToBoolean(Args[3], false);

  sheet := '';
  if Length(Args) > 4 then
    sheet := ArgToString(Args[4]);

  if A1Dialect then
    resStr := GetCellString(r, c, flags)
  else
    resStr := GetCellString_R1C1(r, c, flags);

  if sheet <> '' then resStr := sheet + '!' + resStr;

//  Result := CellResult(resStr);
  Result := StringResult(resStr);
end;

{ COLUMN( [reference] )
  Returns the column number of a cell reference (starting at 1!)
  "reference" is a reference to a cell or range of cells.
  If omitted, it is assumed that the reference is the cell address in which the
  COLUMN function has been entered in. }
procedure fpsCOLUMN(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
begin
  if IsError(Args[0], Result) then
    exit;
  Result := ErrorResult(errArgError);
  if Length(Args) = 0 then
    exit;  // We don't know here which cell contains the formula.
  case Args[0].ResultType of
    rtCell     : Result := IntegerResult(Args[0].ResCol + 1);
    rtCellRange: Result := IntegerResult(Args[0].ResCellRange.Col1 + 1);
    else         Result := ErrorResult(errWrongType);
  end;
end;

procedure fpsHYPERLINK(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
begin
  if IsError(Args[0], Result) then
    exit;
  if (Length(Args) > 1) and IsError(Args[1], Result) then
    exit;

{  if Args[0].ResultType = rtError then
  begin
    Result := ErrorResult(errWrongType);
    exit;
  end;
  if (Length(Args) > 1) and (Args[1].ResultType = rtError) then
  begin
    Result := ErrorResult(errWrongType);
    exit;
  end;
  }
  Result.ResString := ArgToString(Args[0]);
  if Length(Args) > 1 then Result.ResString := Result.ResString + HYPERLINK_SEPARATOR + ArgToString(Args[1]);
  Result.ResultType := rtHyperlink;
end;

{ INDEX(range, row_no [, col_no])
  Searches for a value in an array based on its coordinates.
  When the last parameter, col_no, is omitted, the input range must be a
  1-d vector, either a row or a column.
  Specification of row_no or col_no as zero returns the entire column or row
  within the range. }
procedure fpsINDEX(var Result: TsExpressionResult; const Args: TsExprParameterArray);
var
  rng: TsCellRange3d;
  col: Int64;  // should be cardinal, but using integer due to subtraction
  row: Int64;  // dto.
  book: TsWorkbook;
begin
  if IsError(Args[0], Result) then
    exit;
  if IsError(Args[1], Result) then
    exit;
  if (Length(Args) > 2) and IsError(Args[2], Result) then
    exit;

  Result := ErrorResult(errArgError);
  if Length(Args) < 2 then
    exit;
  if not (Args[0].ResultType = rtCellRange) then
    exit;
  rng := Args[0].ResCellRange;
  if (rng.Sheet1 <> rng.Sheet2) then     // Only range within same sheet allowed.
    exit;

  // 1-d array as column
  if rng.Col1 = rng.Col2 then
  begin
    row := ArgToInt(Args[1]) + rng.Row1 - 1;
    col := rng.Col1;
  end else
  // 1-d array as row
  if rng.Row1 = rng.Row2 then
  begin
    col := ArgToInt(Args[1]) + rng.Col1 - 1;
    row := rng.Row1;
  end else
  // 2-d array
  begin
    if Length(Args) < 3 then
      exit;
    row := ArgToInt(Args[1]) + rng.Row1 - 1;   // The Args are relative to the range
    col := ArgToInt(Args[2]) + rng.Col1 - 1;
    if ArgToInt(Args[1]) = 0 then // entire column within range
    begin
      if (col >= rng.Col1) and (col <= rng.Col2) then
        Result := CellRangeResult(Args[0].Worksheet, rng.Sheet1, rng.Sheet2, rng.Row1, col, rng.Row2, col)
      else
        Result := ErrorResult(errIllegalRef);
      exit;
    end;
    if ArgToInt(Args[2]) = 0 then  // entire row within range
    begin
      if (row >= rng.Row1) and (row <= rng.Row2) then
        Result := CellRangeResult(Args[0].Worksheet, rng.Sheet1, rng.Sheet2, row, rng.Col1, row, rng.Col2)
      else
        Result := ErrorResult(errIllegalRef);
      exit;
    end;
  end;

  // Check whether col/row indices are inside the range
  if (row < rng.Row1) or (row > rng.Row2) or (col < rng.Col1) or (col > rng.Col2) then
  begin
    Result := ErrorResult(errIllegalRef); // #REF! when row/col are outside the range
    exit;
  end;

  book := TsWorksheet(Args[0].Worksheet).Workbook;
  Result := CellResult(row, col);
  Result.Worksheet := Args[0].Worksheet;
  Result.ResSheetIndex := rng.Sheet1;
  Result.ResSheetName := book.GetWorksheetByIndex(rng.Sheet1).Name;
end;

{ INDIRECT(string_reference, [ref_style])
  returns the reference to a cell based on its string representation
  "string_reference": textual representation of a cell reference.
  "ref_style": TRUE (default) indicates that string_reference will be interpreted
  as an A1-style reference. FALSE indicates that string_reference will be
  interpreted as an R1C1-style reference.

  NOTE: ref_style and mixing of A1 and R1C1 notation is not supported. }
procedure fpsINDIRECT(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
begin
  if IsError(Args[0], Result) then
    exit;
  if (Length(Args) > 1) and IsError(Args[1], Result) then
    exit;

  Result := ErrorResult(errArgError);
  if Length(Args) > 0 then
    case Args[0].ResultType of
      rtCell   : Result := CellResult(ArgToString(Args[0]));
      rtString : Result := CellResult(Args[0].ResString);
      rtError  : Result := ErrorResult(Args[0].ResError);
    end;
end;

{ MATCH( value, array, [match_type]
   match_type = 1 (default): The MATCH function will find the largest value
     that is less than or equal to value. You should be sure to sort your
     array in ascending order.
   match_type = 0: The MATCH function will find the first value that is equal to
     value. The array can be sorted in any order.)
   match_type = -1: The MATCH function will find the smallest value that is
     greater than or equal to value. You should be sure to sort your array in
     descending order.
   Return value is the 1-based index in the array.
}
procedure fpsMATCH(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
var
  match_type: Integer;
  searchString: String;
  numSearchValue: Double = 0.0;
  r1, c1, r2, c2: Cardinal;
  r, c: Integer;
  IsCol: Boolean;
  arg: TsExpressionResult;
  sheet: TsWorksheet;
  book: TsWorkbook;
  cell: PCell;

  function Matches(ACell: PCell): Boolean;
  var
    cellval: Double;
    s: String;
  begin
    Result := false;
    if ACell = nil then exit;
    if ACell^.ContentType = cctUTF8String then begin
      s := ACell^.UTF8StringValue;
      if IsWild(searchString, '*?', false) then
        Result := FindPart(searchString, s) > 0
        // NOTE: FindPart currently supports only the wildcard '?'
      else
        Result := CompareStr(s, searchString) = 0;
        //Result := SameStr(s, searchString);
    end else
    begin
      case ACell^.ContentType of
        cctNumber: cellval := ACell^.Numbervalue;
        cctDateTime: cellval := ACell^.DateTimeValue;
        cctBool: cellval := double(ord(ACell^.BoolValue));
        cctError: cellval := double(ord(ACell^.ErrorValue));
        cctEmpty: exit;
      end;
      case match_type of
        1 : Result := cellval <= numSearchValue;
        0 : Result := cellval = numSearchValue;
       -1 : Result := cellval >= numSearchValue;
      end;
    end;
  end;

begin
  Result := ErrorResult(errArgError);

  if Length(Args) > 0 then
    if IsError(Args[0], Result) then
      exit;
  if Length(Args) > 1 then
    if IsError(Args[1], Result) then
      exit;
  if Length(Args) > 2 then
  begin
    if IsError(Args[2], Result) then
      exit;
    match_type := ArgToInt(Args[2])
  end else
    match_type := 1;
  if not ((match_type in [0, 1]) or (match_type = -1)) then
    match_type := 1;

  arg := Args[1];
  if arg.ResultType <> rtCellRange then
    exit;

  if arg.ResCellRange.Sheet1 <> arg.ResCellRange.Sheet2 then
    exit;

  r1 := arg.ResCellRange.Row1;
  r2 := arg.ResCellRange.Row2;
  c1 := arg.ResCellRange.Col1;
  c2 := arg.ResCellRange.Col2;

  if r1=r2 then
    IsCol := false
  else
  if c1=c2 then
    IsCol := true
  else begin
    Result := ErrorResult(errArgError);
    exit;
  end;
  sheet := arg.Worksheet as TsWorksheet;
  book := sheet.Workbook as TsWorkbook;
  sheet := book.GetWorksheetByIndex(arg.ResCellRange.Sheet1);

  case Args[0].ResultType of
    rtString:
      searchString := ArgToString(Args[0]);
    rtCell:
      begin
        cell := ArgToCell(Args[0]);
        if cell = nil then
        begin
          Result := ErrorResult(errArgError);
          exit;
        end;
        case cell^.ContentType of
          cctUTF8String: searchString := cell^.UTF8StringValue;
          cctNumber: numSearchValue := cell^.NumberValue;
          cctDateTime: numSearchValue := cell^.DateTimeValue;
          cctBool: numSearchValue := ord(cell^.BoolValue);
          cctEmpty: begin Result := ErrorResult(errArgError); exit; end;
          cctError: begin Result := ErrorResult(cell^.ErrorValue); exit; end;
        end;
      end;
    else
      numSearchvalue := ArgToFloat(Args[0]);
      if IsNaN(numSearchValue) then begin
        Result := ErrorResult(errWrongType);
        exit;
      end;
  end;

  if IsCol then
  begin
    if match_type = 0 then
    begin
      for r := r1 to r2 do
        if Matches(sheet.Findcell(r, c1)) then
        begin
          Result := IntegerResult(r - integer(r1) + 1);
          exit;
        end;
    end else
    begin
      for r := r2 downto r1 do
        if Matches(sheet.FindCell(r, c1)) then begin
          Result := IntegerResult(r - integer(r1) + 1);
          exit;
        end;
    end;
  end else
  begin
    if match_type = 0 then
    begin
      for c := c1 to c2 do
        if Matches(sheet.Findcell(r1, c)) then
        begin
          Result := IntegerResult(c - Integer(c1) + 1);
          exit;
        end;
    end else
    begin
      for c := c2 downto c1 do
        if Matches(sheet.FindCell(r1, c)) then begin
          Result := IntegerResult(c - Integer(c1) + 1);
          exit;
      end;
    end;
  end;

  // If the procedure gets here, not match has been found --> return error #N/A
end;

{ ROW( [reference] )
  Returns the row number of a cell reference (starting at 1!)
  "reference" is a reference to a cell or range of cells.
  If omitted, it is assumed that the reference is the cell address in which the
  ROW function has been entered in. }
procedure fpsROW(var Result: TsExpressionResult;
  const Args: TsExprParameterArray);
begin
  Result := ErrorResult(errArgError);
  if Length(Args) = 0 then
    exit;  // We don't know here which cell contains the formula.
  if IsError(Args[0], Result) then
    exit;
  case Args[0].ResultType of
    rtCell     : Result := IntegerResult(Args[0].ResRow + 1);
    rtCellRange: Result := IntegerResult(Args[0].ResCellRange.Row1 + 1);
    else         Result := ErrorResult(errWrongType);
  end;
end;


{------------------------------------------------------------------------------}
{   Registration                                                               }
{------------------------------------------------------------------------------}

{@@ Registers the standard built-in functions. Called automatically. }
procedure RegisterStdBuiltins(AManager : TComponent);
var
  cat: TsBuiltInExprCategory;
begin
  with AManager as TsBuiltInExpressionManager do
  begin
    // Math functions
    cat := bcMath;
    AddFunction(cat, 'ABS',       'F', 'F',    INT_EXCEL_SHEET_FUNC_ABS,        @fpsABS);
    AddFunction(cat, 'ACOS',      'F', 'F',    INT_EXCEL_SHEET_FUNC_ACOS,       @fpsACOS);
    AddFunction(cat, 'ACOSH',     'F', 'F',    INT_EXCEL_SHEET_FUNC_ACOSH,      @fpsACOSH);
    AddFunction(cat, 'ASIN',      'F', 'F',    INT_EXCEL_SHEET_FUNC_ASIN,       @fpsASIN);
    AddFunction(cat, 'ASINH',     'F', 'F',    INT_EXCEL_SHEET_FUNC_ASINH,      @fpsASINH);
    AddFunction(cat, 'ATAN',      'F', 'F',    INT_EXCEL_SHEET_FUNC_ATAN,       @fpsATAN);
    AddFunction(cat, 'ATANH',     'F', 'F',    INT_EXCEL_SHEET_FUNC_ATANH,      @fpsATANH);
    AddFunction(cat, 'CEILING',   'F', 'FF',   INT_EXCEL_SHEET_FUNC_CEILING,    @fpsCEILING);
    AddFunction(cat, 'COS',       'F', 'F',    INT_EXCEL_SHEET_FUNC_COS,        @fpsCOS);
    AddFunction(cat, 'COSH',      'F', 'F',    INT_EXCEL_SHEET_FUNC_COSH,       @fpsCOSH);
    AddFunction(cat, 'DEGREES',   'F', 'F',    INT_EXCEL_SHEET_FUNC_DEGREES,    @fpsDEGREES);
    AddFunction(cat, 'EVEN',      'I', 'F',    INT_EXCEL_SHEET_FUNC_EVEN,       @fpsEVEN);
    AddFunction(cat, 'EXP',       'F', 'F',    INT_EXCEL_SHEET_FUNC_EXP,        @fpsEXP);
    AddFunction(cat, 'FACT',      'F', 'I',    INT_EXCEL_SHEET_FUNC_FACT,       @fpsFACT);
    AddFunction(cat, 'FLOOR',     'F', 'FF',   INT_EXCEL_SHEET_FUNC_FLOOR,      @fpsFLOOR);
    AddFunction(cat, 'INT',       'I', 'F',    INT_EXCEL_SHEET_FUNC_INT,        @fpsINT);
    AddFunction(cat, 'LN',        'F', 'F',    INT_EXCEL_SHEET_FUNC_LN,         @fpsLN);
    AddFunction(cat, 'LOG',       'F', 'Ff',   INT_EXCEL_SHEET_FUNC_LOG,        @fpsLOG);
    AddFunction(cat, 'LOG10',     'F', 'F',    INT_EXCEL_SHEET_FUNC_LOG10,      @fpsLOG10);
    AddFunction(cat, 'MOD',       'I', 'II',   INT_EXCEL_SHEET_FUNC_MOD,        @fpsMOD);
    AddFunction(cat, 'ODD',       'I', 'F',    INT_EXCEL_SHEET_FUNC_ODD,        @fpsODD);
    AddFunction(cat, 'PI',        'F', '',     INT_EXCEL_SHEET_FUNC_PI,         @fpsPI);
    AddFunction(cat, 'POWER',     'F', 'FF',   INT_EXCEL_SHEET_FUNC_POWER,      @fpsPOWER);
    AddFunction(cat, 'RADIANS',   'F', 'F',    INT_EXCEL_SHEET_FUNC_RADIANS,    @fpsRADIANS);
    AddFunction(cat, 'RAND',      'F', '',     INT_EXCEL_SHEET_FUNC_RAND,       @fpsRAND);
    AddFunction(cat, 'ROUND',     'F', 'FF',   INT_EXCEL_SHEET_FUNC_ROUND,      @fpsROUND);
    AddFunction(cat, 'ROUNDDOWN', 'F', 'FF',   INT_EXCEL_SHEET_FUNC_ROUNDDOWN,  @fpsROUNDDOWN);
    AddFunction(cat, 'ROUNDUP',   'F', 'FF',   INT_EXCEL_SHEET_FUNC_ROUNDUP,    @fpsROUNDUP);
    AddFunction(cat, 'SIGN',      'F', 'F',    INT_EXCEL_SHEET_FUNC_SIGN,       @fpsSIGN);
    AddFunction(cat, 'SIN',       'F', 'F',    INT_EXCEL_SHEET_FUNC_SIN,        @fpsSIN);
    AddFunction(cat, 'SINH',      'F', 'F',    INT_EXCEL_SHEET_FUNC_SINH,       @fpsSINH);
    AddFunction(cat, 'SQRT',      'F', 'F',    INT_EXCEL_SHEET_FUNC_SQRT,       @fpsSQRT);
    AddFunction(cat, 'TAN',       'F', 'F',    INT_EXCEL_SHEET_FUNC_TAN,        @fpsTAN);
    AddFunction(cat, 'TANH',      'F', 'F',    INT_EXCEL_SHEET_FUNC_TANH,       @fpsTANH);

    // Date/time
    cat := bcDateTime;
    AddFunction(cat, 'DATE',      'D', 'III',  INT_EXCEL_SHEET_FUNC_DATE,       @fpsDATE);
    AddFunction(cat, 'DATEDIF',   'F', 'DDS',  INT_EXCEL_SHEET_FUNC_DATEDIF,    @fpsDATEDIF);
    AddFunction(cat, 'DATEVALUE', 'D', 'S',    INT_EXCEL_SHEET_FUNC_DATEVALUE,  @fpsDATEVALUE);
    AddFunction(cat, 'DAY',       'I', '?',    INT_EXCEL_SHEET_FUNC_DAY,        @fpsDAY);
    AddFunction(cat, 'HOUR',      'I', '?',    INT_EXCEL_SHEET_FUNC_HOUR,       @fpsHOUR);
    AddFunction(cat, 'MINUTE',    'I', '?',    INT_EXCEL_SHEET_FUNC_MINUTE,     @fpsMINUTE);
    AddFunction(cat, 'MONTH',     'I', '?',    INT_EXCEL_SHEET_FUNC_MONTH,      @fpsMONTH);
    AddFunction(cat, 'NOW',       'D', '',     INT_EXCEL_SHEET_FUNC_NOW,        @fpsNOW);
    AddFunction(cat, 'SECOND',    'I', '?',    INT_EXCEL_SHEET_FUNC_SECOND,     @fpsSECOND);
    AddFunction(cat, 'TIME' ,     'D', 'III',  INT_EXCEL_SHEET_FUNC_TIME,       @fpsTIME);
    AddFunction(cat, 'TIMEVALUE', 'D', 'S',    INT_EXCEL_SHEET_FUNC_TIMEVALUE,  @fpsTIMEVALUE);
    AddFunction(cat, 'TODAY',     'D', '',     INT_EXCEL_SHEET_FUNC_TODAY,      @fpsTODAY);
    AddFunction(cat, 'WEEKDAY',   'I', '?i',   INT_EXCEL_SHEET_FUNC_WEEKDAY,    @fpsWEEKDAY);
    AddFunction(cat, 'YEAR',      'I', '?',    INT_EXCEL_SHEET_FUNC_YEAR,       @fpsYEAR);

    // Strings
    cat := bcStrings;
    AddFunction(cat, 'CHAR',      'S', 'I',    INT_EXCEL_SHEET_FUNC_CHAR,       @fpsCHAR);
    AddFunction(cat, 'CODE',      'I', 'S',    INT_EXCEL_SHEET_FUNC_CODE,       @fpsCODE);
    AddFunction(cat, 'CONCATENATE','S','S+',   INT_EXCEL_SHEET_FUNC_CONCATENATE,@fpsCONCATENATE);
    AddFunction(cat, 'EXACT',     'B', 'SS',   INT_EXCEL_SHEET_FUNC_EXACT,      @fpsEXACT);
    AddFunction(cat, 'LEFT',      'S', 'Si',   INT_EXCEL_SHEET_FUNC_LEFT,       @fpsLEFT);
    AddFunction(cat, 'LEN',       'I', 'S',    INT_EXCEL_SHEET_FUNC_LEN,        @fpsLEN);
    AddFunction(cat, 'LOWER',     'S', 'S',    INT_EXCEL_SHEET_FUNC_LOWER,      @fpsLOWER);
    AddFunction(cat, 'MID',       'S', 'SII',  INT_EXCEL_SHEET_FUNC_MID,        @fpsMID);
    AddFunction(cat, 'REPLACE',   'S', 'SIIS', INT_EXCEL_SHEET_FUNC_REPLACE,    @fpsREPLACE);
    AddFunction(cat, 'REPT',      'S', 'SI',   INT_EXCEL_SHEET_FUNC_REPT,       @fpsREPT);
    AddFunction(cat, 'RIGHT',     'S', 'Si',   INT_EXCEL_SHEET_FUNC_RIGHT,      @fpsRIGHT);
    AddFunction(cat, 'SUBSTITUTE','S', 'SSSi', INT_EXCEL_SHEET_FUNC_SUBSTITUTE, @fpsSUBSTITUTE);
    AddFunction(cat, 'TEXT',      'S', '?S',   INT_EXCEL_SHEET_FUNC_TEXT,       @fpsTEXT);
    AddFunction(cat, 'TRIM',      'S', 'S',    INT_EXCEL_SHEET_FUNC_TRIM,       @fpsTRIM);
    AddFunction(cat, 'UPPER',     'S', 'S',    INT_EXCEL_SHEET_FUNC_UPPER,      @fpsUPPER);
    AddFunction(cat, 'VALUE',     'F', 'S',    INT_EXCEL_SHEET_FUNC_VALUE,      @fpsVALUE);

    // Logical
    cat := bcLogical;
    AddFunction(cat, 'AND',       'B', 'B+',   INT_EXCEL_SHEET_FUNC_AND,        @fpsAND);
    AddFunction(cat, 'FALSE',     'B', '',     INT_EXCEL_SHEET_FUNC_FALSE,      @fpsFALSE);
    AddFunction(cat, 'IF',        'B', 'B?+',  INT_EXCEL_SHEET_FUNC_IF,         @fpsIF);
//  AddFunction(cat, 'IFS',       'B', 'B?+',  INT_EXCEL_SHEET_FUNC_UNKNOWN,    @fpsIFS);
//    AddFunction(cat, 'COM.MICROSOFT.IFS', 'B', 'B?+',  INT_EXCEL_SHEET_FUNC_UNKNOWN,    @fpsIFS);
    AddFunction(cat, '_xlfn.IFS', 'B', 'B?+',  INT_EXCEL_SHEET_FUNC_UNKNOWN,    @fpsIFS);

    AddFunction(cat, 'NOT',       'B', 'B',    INT_EXCEL_SHEET_FUNC_NOT,        @fpsNOT);
    AddFunction(cat, 'OR',        'B', 'B+',   INT_EXCEL_SHEET_FUNC_OR,         @fpsOR);
    AddFunction(cat, 'TRUE',      'B', '',     INT_EXCEL_SHEET_FUNC_TRUE ,      @fpsTRUE);

    // Statistical
    cat := bcStatistics;
    AddFunction(cat, 'AVEDEV',    'F', 'F+',   INT_EXCEL_SHEET_FUNC_AVEDEV,     @fpsAVEDEV);
    AddFunction(cat, 'AVERAGE',   'F', 'F+',   INT_EXCEL_SHEET_FUNC_AVERAGE,    @fpsAVERAGE);
    AddFunction(cat, 'AVERAGEIF', 'F', 'R?r',  INT_EXCEL_SHEET_FUNC_NOT_BIFF,   @fpsAVERAGEIF);
    AddFunction(cat, 'AVERAGEIFS','F', 'R?r+', INT_EXCEL_SHEET_FUNC_NOT_BIFF,   @fpsAVERAGEIFS);
    AddFunction(cat, 'COUNT',     'I', '?+',   INT_EXCEL_SHEET_FUNC_COUNT,      @fpsCOUNT);
    AddFunction(cat, 'COUNTA',    'I', '?+',   INT_EXCEL_SHEET_FUNC_COUNTA,     @fpsCOUNTA);
    AddFunction(cat, 'COUNTBLANK','I', 'R',    INT_EXCEL_SHEET_FUNC_COUNTBLANK, @fpsCOUNTBLANK);
    AddFunction(cat, 'COUNTIF',   'I', 'R?',   INT_EXCEL_SHEET_FUNC_COUNTIF,    @fpsCOUNTIF);
    AddFunction(cat, 'COUNTIFS',  'I', 'R?r+', INT_EXCEL_SHEET_FUNC_NOT_BIFF,   @fpsCOUNTIFS);
    AddFunction(cat, 'MAX',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_MAX,        @fpsMAX);
    AddFunction(cat, 'MIN',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_MIN,        @fpsMIN);
    AddFunction(cat, 'PRODUCT',   'F', 'F+',   INT_EXCEL_SHEET_FUNC_PRODUCT,    @fpsPRODUCT);
    AddFunction(cat, 'STDEV',     'F', 'F+',   INT_EXCEL_SHEET_FUNC_STDEV,      @fpsSTDEV);
    AddFunction(cat, 'STDEVP',    'F', 'F+',   INT_EXCEL_SHEET_FUNC_STDEVP,     @fpsSTDEVP);
    AddFunction(cat, 'SUM',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_SUM,        @fpsSUM);
    AddFunction(cat, 'SUMIF',     'F', 'R?r',  INT_EXCEL_SHEET_FUNC_SUMIF,      @fpsSUMIF);
    AddFunction(cat, 'SUMIFS',    'F', 'R?r+', INT_EXCEL_SHEET_FUNC_NOT_BIFF,   @fpsSUMIFS);
    AddFunction(cat, 'SUMSQ',     'F', 'F+',   INT_EXCEL_SHEET_FUNC_SUMSQ,      @fpsSUMSQ);
    AddFunction(cat, 'VAR',       'F', 'F+',   INT_EXCEL_SHEET_FUNC_VAR,        @fpsVAR);
    AddFunction(cat, 'VARP',      'F', 'F+',   INT_EXCEL_SHEET_FUNC_VARP,       @fpsVARP);

    // Info functions
    cat := bcInfo;
    //AddFunction(cat, 'CELL',      '?', 'Sr',   INT_EXCEL_SHEET_FUNC_CELL,       @fpsCELL);
    AddFunction(cat, 'ERROR.TYPE','I', '?',    INT_EXCEL_SHEET_FUNC_ERRORTYPE,  @fpsERRORTYPE);
    AddFunction(cat, 'ISBLANK',   'B', '?',    INT_EXCEL_SHEET_FUNC_ISBLANK,    @fpsISBLANK);
    AddFunction(cat, 'ISERR',     'B', '?',    INT_EXCEL_SHEET_FUNC_ISERR,      @fpsISERR);
    AddFunction(cat, 'ISERROR',   'B', '?',    INT_EXCEL_SHEET_FUNC_ISERROR,    @fpsISERROR);
    AddFunction(cat, 'ISLOGICAL', 'B', '?',    INT_EXCEL_SHEET_FUNC_ISLOGICAL,  @fpsISLOGICAL);
    AddFunction(cat, 'ISNA',      'B', '?',    INT_EXCEL_SHEET_FUNC_ISNA,       @fpsISNA);
    AddFunction(cat, 'ISNONTEXT', 'B', '?',    INT_EXCEL_SHEET_FUNC_ISNONTEXT,  @fpsISNONTEXT);
    AddFunction(cat, 'ISNUMBER',  'B', '?',    INT_EXCEL_SHEET_FUNC_ISNUMBER,   @fpsISNUMBER);
    AddFunction(cat, 'ISREF',     'B', '?',    INT_EXCEL_SHEET_FUNC_ISREF,      @fpsISREF);
    AddFunction(cat, 'ISTEXT',    'B', '?',    INT_EXCEL_SHEET_FUNC_ISTEXT,     @fpsISTEXT);

    AddFunction(cat, 'IFERROR',   '?', '?S',   INT_EXCEL_SHEET_FUNC_UNKNOWN,    @fpsIFERROR);   // not supported by .xls

    // Lookup / reference functions
    cat := bcLookup;
    AddFunction(cat, 'ADDRESS',   'S', 'IIibs',INT_EXCEL_SHEET_FUNC_ADDRESS,    @fpsADDRESS);
    AddFunction(cat, 'COLUMN',    'I', 'r',    INT_EXCEL_SHEET_FUNC_COLUMN,     @fpsCOLUMN);
    AddFunction(cat, 'HYPERLINK', 'S', 'Ss',   INT_EXCEL_SHEET_FUNC_HYPERLINK,  @fpsHYPERLINK);
    AddFunction(cat, 'INDEX',     'C', 'RIi',  INT_EXCEL_SHEET_FUNC_INDEX,      @fpsINDEX);
    AddFunction(cat, 'INDIRECT',  'C', 'Sb',   INT_EXCEL_SHEET_FUNC_INDIRECT,   @fpsINDIRECT);
    AddFunction(cat, 'MATCH',     'I', 'SRi',  INT_EXCEL_SHEET_FUNC_MATCH,      @fpsMATCH);
    AddFunction(cat, 'ROW',       'I', 'r',    INT_EXCEL_SHEET_FUNC_ROW,        @fpsROW);

    (*
    AddFunction(cat, 'COLUMN',    'I', 'R',    INT_EXCEL_SHEET_FUNC_COLUMN,     @fpsCOLUMN);
                  *)

  end;
end;


{ Lookup / reference functions }

(*
function fpsCOLUMNS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ COLUMNS( [reference] )
  returns the number of column in a cell reference. }
var
  arg: TsArgument;
begin
  Unused(NumArgs);
  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumberArg(1);
    atCellRange: Result := CreateNumberArg(arg.LastCol - arg.FirstCol + 1);
    else         Result := CreateErrorArg(errWrongType);
  end;
end;


function fpsROWS(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{ ROWS( [reference] )
  returns the number of rows in a cell reference. }
var
  arg: TsArgument;
begin
  Unused(NumArgs);
  arg := Args.Pop;
  case arg.ArgumentType of
    atCell     : Result := CreateNumberArg(1);
    atCellRange: Result := CreateNumberArg(arg.LastRow - arg.FirstRow + 1);
    else         Result := CreateErrorArg(errWrongType);
  end;
end;

*)

(*
function fpsINFO(Args: TsArgumentStack; NumArgs: Integer): TsArgument;
{  INFO( type )
   returns information about the operating environment.
   type can be one of the following values:
     + "directory"    Path of the current directory.
     + "numfile"      Number of active worksheets.
     - "origin"       The cell that is in the top, left-most cell visible in the current Excel spreadsheet.
     - "osversion"    Operating system version.
     - "recalc"       Returns the recalculation mode - either Automatic or Manual.
     - "release"      Version of Excel that you are running.
     - "system"       Name of the operating environment.
   ONLY THOSE MARKED BY "+" ARE SUPPORTED! }
var
  arg: TsArgument;
  workbook: TsWorkbook;
  s: String;
begin
  Unused(NumArgs);
  arg := Args.Pop;
  if arg.ArgumentType <> atString then
    Result := CreateErrorArg(errWrongType)
  else begin
    s := Lowercase(arg.StringValue);
    workbook := arg.Worksheet.Workbook;
    if s = 'directory' then
      Result := CreateStringArg(ExtractFilePath(workbook.FileName))
    else
    if s = 'numfile' then
      Result := CreateNumberArg(workbook.GetWorksheetCount)
    else
      Result := CreateErrorArg(errFormulaNotSupported);
  end;
end;

*)


end.
