{
    This file is part of the Free Component Library (FCL)
    Copyright (c) 2008 Michael Van Canneyt.

    Expression parser, supports variables, functions and
    float/integer/string/boolean/datetime operations.

    See the file COPYING.FPC, included in this distribution,
    for details about the copyright.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

--------------------------------------------------------------------------------

    Modified for integration into fpspreadsheet by Werner Pamler:
    - Original file name: fpexprpars.pp
    - Rename identifiers to avoid naming conflicts with the original
    - TsExpressionParser and TsBuiltinExpressionManager are not components
      any more
    - TsExpressionParser is created with the worksheet as a parameter.
    - add new TExprNode classes:
      - TsCellExprNode for references to cells
      - TsCellRangeExprNode for references to cell ranges
      - TsPercentExprNode and token "%" to handle Excel's percent operation
      - TsParenthesisExprNode to handle the parenthesis token in RPN formulas
      - TsConcatExprNode and token "&" to handle string concatenation
      - TsUPlusExprNode for unary plus symbol
    - remove and modifiy built-in function such that the parser is compatible
      with Excel syntax (and Open/LibreOffice - which is the same).
    - use double quotes for strings (instead of single quotes)
    - add boolean constants "TRUE" and "FALSE".
    - add property RPNFormula to interface the parser to RPN formulas of xls files.
    - accept funtions with zero parameters
    - generalize scanner and parser to allow localized decimal and list separators
    - add to spreadsheet format to parser to take account of formula "dialect"
      (see OpenDocument using [] around cell addresses)

 ******************************************************************************}

// To do:
// Remove exceptions, use error message strings instead
// Cell reference not working (--> formula CELL!)
// Keep spaces in formula

{$mode objfpc}
{$h+}
unit fpsExprParser;

interface

uses
  Classes, SysUtils, contnrs, fpstypes, fpspreadsheet, fpsrpn;

type
  { Tokens }

  TsTokenType = (
    ttCell, ttCellRange, ttSheetName,
    ttNumber, ttString, ttIdentifier,
    ttPlus, ttMinus, ttMul, ttDiv, ttConcat, ttPercent, ttPower, ttLeft, ttRight,
    ttLessThan, ttLargerThan, ttEqual, ttNotEqual, ttLessThanEqual, ttLargerThanEqual,
    ttListSep, ttTrue, ttFalse, ttMissingArg, ttError, ttEOF
  );

  TsExprFloat = Double;
  TsExprFloatArray = array of TsExprFloat;

const
  ttDelimiters = [
    ttPlus, ttMinus, ttMul, ttDiv, ttLeft, ttRight, ttLessThan, ttLargerThan,
    ttEqual, ttNotEqual, ttLessThanEqual, ttLargerThanEqual
  ];

  ttComparisons = [
    ttLargerThan, ttLessThan, ttLargerThanEqual, ttLessThanEqual, ttEqual, ttNotEqual
  ];

type
  TsExpressionParser = class;
  TsBuiltInExpressionManager = class;

  TsResultType = (rtEmpty, rtBoolean, rtInteger, rtFloat, rtDateTime, rtString,
    rtCell, rtCellRange, rtHyperlink, rtError, rtMissingArg, rtAny);
  TsResultTypes = set of TsResultType;

  TsExpressionResult = record
    Worksheet       : TsWorksheet;
    ResString       : String;
    case ResultType : TsResultType of
      rtEmpty       : ();
      rtError       : (ResError       : TsErrorValue);
      rtBoolean     : (ResBoolean     : Boolean);
      rtInteger     : (ResInteger     : Int64);
      rtFloat       : (ResFloat       : TsExprFloat);
      rtDateTime    : (ResDateTime    : TDatetime);
      rtCell        : (ResRow, ResCol : Cardinal);
      rtCellRange   : (ResCellRange   : TsCellRange);
      rtHyperlink   : ();
      rtString      : ();
  end;
  PsExpressionResult = ^TsExpressionResult;
  TsExprParameterArray = array of TsExpressionResult;

  { TsExprNode }
  TsExprNode = class(TObject)
  private
    FParser: TsExpressionParser;
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); virtual; abstract;
    function HasError(out AResult: TsExpressionResult): boolean; virtual;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; virtual; abstract;
    function AsString: string; virtual; abstract;
    procedure Check; virtual; //abstract;
    function NodeType: TsResultType; virtual; abstract;
    function NodeValue: TsExpressionResult;
    property Parser: TsExpressionParser read FParser;
  end;

  TsExprArgumentArray = array of TsExprNode;

  { TsBinaryOperationExprNode }
  TsBinaryOperationExprNode = class(TsExprNode)
  private
    FLeft: TsExprNode;
    FRight: TsExprNode;
  protected
    function HasError(out AResult: TsExpressionResult): Boolean; override;
  public
    constructor Create(AParser: TsExpressionParser; ALeft, ARight: TsExprNode);
    destructor Destroy; override;
    property Left: TsExprNode read FLeft;
    property Right: TsExprNode read FRight;
  end;
  TsBinaryOperationExprNodeClass = class of TsBinaryOperationExprNode;

  { TsBooleanOperationExprNode }
  TsBooleanOperationExprNode = class(TsBinaryOperationExprNode)
  public
    function NodeType: TsResultType; override;
  end;

  { TsBooleanResultExprNode }
  TsBooleanResultExprNode = class(TsBinaryOperationExprNode)
  public
    function NodeType: TsResultType; override;
  end;
  TsBooleanResultExprNodeClass = class of TsBooleanResultExprNode;

  { TsEqualExprNode }
  TsEqualExprNode = class(TsBooleanResultExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsNotEqualExprNode }
  TsNotEqualExprNode = class(TsEqualExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsOrderingExprNode }
  TsOrderingExprNode = class(TsBooleanResultExprNode);

  { TsLessExprNode }
  TsLessExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterExprNode }
  TsGreaterExprNode = class(TsOrderingExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsLessEqualExprNode }
  TsLessEqualExprNode = class(TsGreaterExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsGreaterEqualExprNode }
  TsGreaterEqualExprNode = class(TsLessExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;

  { TsConcatExprNode }
  TsConcatExprNode = class(TsBinaryOperationExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    function NodeType: TsResultType; override;
  end;

  { TsMathOperationExprNode }
  TsMathOperationExprNode = class(TsBinaryOperationExprNode)
  public
    function NodeType: TsResultType; override;
  end;

  { TsAddExprNode }
  TsAddExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsSubtractExprNode }
  TsSubtractExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsMultiplyExprNode }
  TsMultiplyExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
  end;

  { TsDivideExprNode }
  TsDivideExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    function NodeType: TsResultType; override;
  end;

  { TsPowerExprNode }
  TsPowerExprNode = class(TsMathOperationExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string ; override;
    function NodeType: TsResultType; override;
  end;

  { TsUnaryOperationExprNode }
  TsUnaryOperationExprNode = class(TsExprNode)
  private
    FOperand: TsExprNode;
  public
    constructor Create(AParser: TsExpressionParser; AOperand: TsExprNode);
    procedure Check; override;
    destructor Destroy; override;
    property Operand: TsExprNode read FOperand;
  end;
             (*
  { TsNotExprNode }
  TsNotExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;
              *)
  (*
  { TsConvertExprNode }
  TsConvertExprNode = class(TsUnaryOperationExprNode)
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
  end;

  { TsConvertToIntExprNode }
  TsConvertToIntExprNode = class(TsConvertExprNode)
  public
//    procedure Check; override;
  end;

  { TsIntToFloatExprNode }
  TsIntToFloatExprNode = class(TsConvertToIntExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function NodeType: TsResultType; override;
  end;

  { TsIntToDateTimeExprNode }
  TsIntToDateTimeExprNode = class(TsConvertToIntExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function NodeType: TsResultType; override;
  end;

  { TsFloatToDateTimeExprNode }
  TsFloatToDateTimeExprNode = class(TsConvertExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
//    procedure Check; override;
    function NodeType: TsResultType; override;
  end;
                     *)
  { TsUPlusExprNode }
  TsUPlusExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
//    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsUMinusExprNode }
  TsUMinusExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
//    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsPercentExprNode }
  TsPercentExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
  end;

  { TsParenthesisExprNode }
  TsParenthesisExprNode = class(TsUnaryOperationExprNode)
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    function NodeType: TsResultType; override;
  end;

  { TsConstExprNode }
  TsConstExprNode = class(TsExprNode)
  private
    FValue: TsExpressionResult;
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    constructor CreateString(AParser: TsExpressionParser; AValue: String);
    constructor CreateInteger(AParser: TsExpressionParser; AValue: Int64);
    constructor CreateDateTime(AParser: TsExpressionParser; AValue: TDateTime);
    constructor CreateFloat(AParser: TsExpressionParser; AValue: TsExprFloat);
    constructor CreateBoolean(AParser: TsExpressionParser; AValue: Boolean);
    constructor CreateError(AParser: TsExpressionParser; AValue: TsErrorValue); overload;
    constructor CreateError(AParser: TsExpressionParser; AValue: String); overload;
    function AsString: string; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function NodeType : TsResultType; override;
    // For inspection
    property ConstValue: TsExpressionResult read FValue;
  end;

  { TsMissingArgExprNode }
  TsMissingArgExprNode = class(TsExprNode)
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    function AsString: String; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function NodeType: TsResultType; override;
  end;

  TsExprIdentifierType = (itVariable, itFunctionCallBack, itFunctionHandler);

  TsExprFunctionCallBack = procedure (var Result: TsExpressionResult;
    const Args: TsExprParameterArray);

  TsExprFunctionEvent = procedure (var Result: TsExpressionResult;
    const Args: TsExprParameterArray) of object;

  { TsExprIdentifierDef }
  TsExprIdentifierDef = class(TCollectionItem)
  private
    FStringValue: String;
    FValue: TsExpressionResult;
    FArgumentTypes: String;
    FIDType: TsExprIdentifierType;
    FName: ShortString;
    FExcelCode: Integer;
    FVariableArgumentCount: Boolean;
    FOnGetValue: TsExprFunctionEvent;
    FOnGetValueCB: TsExprFunctionCallBack;
    function GetAsBoolean: Boolean;
    function GetAsDateTime: TDateTime;
    function GetAsFloat: TsExprFloat;
    function GetAsInteger: Int64;
    function GetAsString: String;
    function GetResultType: TsResultType;
    function GetValue: String;
    procedure SetArgumentTypes(const AValue: String);
    procedure SetAsBoolean(const AValue: Boolean);
    procedure SetAsDateTime(const AValue: TDateTime);
    procedure SetAsFloat(const AValue: TsExprFloat);
    procedure SetAsInteger(const AValue: Int64);
    procedure SetAsString(const AValue: String);
    procedure SetName(const AValue: ShortString);
    procedure SetResultType(const AValue: TsResultType);
    procedure SetValue(const AValue: String);
  protected
    procedure CheckResultType(const AType: TsResultType);
    procedure CheckVariable;
    function GetFormatSettings: TFormatSettings;
  public
    function ArgumentCount: Integer;
    procedure Assign(Source: TPersistent); override;
    property AsFloat: TsExprFloat Read GetAsFloat Write SetAsFloat;
    property AsInteger: Int64 Read GetAsInteger Write SetAsInteger;
    property AsString: String Read GetAsString Write SetAsString;
    property AsBoolean: Boolean Read GetAsBoolean Write SetAsBoolean;
    property AsDateTime: TDateTime Read GetAsDateTime Write SetAsDateTime;
    function HasFixedArgumentCount: Boolean;
    function IsOptionalArgument(AIndex: Integer): Boolean;
    property OnGetFunctionValueCallBack: TsExprFunctionCallBack read FOnGetValueCB write FOnGetValueCB;
  published
    property IdentifierType: TsExprIdentifierType read FIDType write FIDType;
    property Name: ShortString read FName write SetName;
    property Value: String read GetValue write SetValue;
    property ParameterTypes: String read FArgumentTypes write SetArgumentTypes;
    property ResultType: TsResultType read GetResultType write SetResultType;
    property ExcelCode: Integer read FExcelCode write FExcelCode;
    property VariableArgumentCount: Boolean read FVariableArgumentCount write FVariableArgumentCount;
    property OnGetFunctionValue: TsExprFunctionEvent read FOnGetValue write FOnGetValue;
  end;

  TsBuiltInExprCategory = (bcMath, bcStatistics, bcStrings, bcLogical, bcDateTime,
    bcLookup, bcInfo, bcUser);

  TsBuiltInExprCategories = set of TsBuiltInExprCategory;

  { TsBuiltInExprIdentifierDef }
  TsBuiltInExprIdentifierDef = class(TsExprIdentifierDef)
  private
    FCategory: TsBuiltInExprCategory;
  public
    procedure Assign(Source: TPersistent); override;
  published
    property Category: TsBuiltInExprCategory read FCategory write FCategory;
  end;

  { TsExprIdentifierDefs }
  TsExprIdentifierDefs = class(TCollection)
  private
    FParser: TsExpressionParser;
    function GetI(AIndex: Integer): TsExprIdentifierDef;
    procedure SetI(AIndex: Integer; const AValue: TsExprIdentifierDef);
  protected
    procedure Update(Item: TCollectionItem); override;
    property Parser: TsExpressionParser read FParser;
  public
    function FindIdentifier(const AName: ShortString): TsExprIdentifierDef;
    function IdentifierByExcelCode(const AExcelCode: Integer): TsExprIdentifierDef;
    function IdentifierByName(const AName: ShortString): TsExprIdentifierDef;
    function IndexOfIdentifier(const AName: ShortString): Integer; overload;
    function IndexOfIdentifier(const AExcelCode: Integer): Integer; overload;
    function AddVariable(const AName: ShortString; AResultType: TsResultType;
      AValue: String): TsExprIdentifierDef;
    function AddBooleanVariable(const AName: ShortString;
      AValue: Boolean): TsExprIdentifierDef;
    function AddIntegerVariable(const AName: ShortString;
      AValue: Integer): TsExprIdentifierDef;
    function AddFloatVariable(const AName: ShortString;
      AValue: TsExprFloat): TsExprIdentifierDef;
    function AddStringVariable(const AName: ShortString;
      AValue: String): TsExprIdentifierDef;
    function AddDateTimeVariable(const AName: ShortString;
      AValue: TDateTime): TsExprIdentifierDef;
    function AddFunction(const AName: ShortString; const AResultType: Char;
      const AParamTypes: String; const AExcelCode: Integer;
      ACallBack: TsExprFunctionCallBack): TsExprIdentifierDef;
    function AddFunction(const AName: ShortString; const AResultType: Char;
      const AParamTypes: String; const AExcelCode: Integer;
      ACallBack: TsExprFunctionEvent): TsExprIdentifierDef;
    property Identifiers[AIndex: Integer]: TsExprIdentifierDef read GetI write SetI; default;
  end;

  { TsIdentifierExprNode }
  TsIdentifierExprNode = class(TsExprNode)
  private
    FID: TsExprIdentifierDef;
    PResult: PsExpressionResult;
    FResultType: TsResultType;
  protected
    procedure GetNodeValue(out AResult: TsExpressionResult); override;
  public
    constructor CreateIdentifier(AParser: TsExpressionParser; AID: TsExprIdentifierDef);
    function NodeType: TsResultType; override;
    property Identifier: TsExprIdentifierDef read FID;
  end;

  { TsVariableExprNode }
  TsVariableExprNode = class(TsIdentifierExprNode)
  public
    function AsString: string; override;
    Function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
  end;

  { TsFunctionExprNode }
  TsFunctionExprNode = class(TsIdentifierExprNode)
  private
    FArgumentNodes: TsExprArgumentArray;
    FargumentParams: TsExprParameterArray;
  protected
    procedure CalcParams;
  public
    constructor CreateFunction(AParser: TsExpressionParser;
      AID: TsExprIdentifierDef; const Args: TsExprArgumentArray); virtual;
    destructor Destroy; override;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    property ArgumentNodes: TsExprArgumentArray read FArgumentNodes;
    property ArgumentParams: TsExprParameterArray read FArgumentParams;
  end;

  { TsFunctionCallBackExprNode }
  TsFunctionCallBackExprNode = class(TsFunctionExprNode)
  private
    FCallBack: TsExprFunctionCallBack;
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor CreateFunction(AParser: TsExpressionParser;
      AID: TsExprIdentifierDef; const Args: TsExprArgumentArray); override;
    property CallBack: TsExprFunctionCallBack read FCallBack;
  end;

  { TFPFunctionEventHandlerExprNode }
  TFPFunctionEventHandlerExprNode = class(TsFunctionExprNode)
  private
    FCallBack: TsExprFunctionEvent;
  protected
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor CreateFunction(AParser: TsExpressionParser;
      AID: TsExprIdentifierDef; const Args: TsExprArgumentArray); override;
    property CallBack: TsExprFunctionEvent read FCallBack;
  end;

  TsSheetNameExprNode = class(TsExprNode)
  private
    FSheetName: String;
  public
    constructor Create(AParser: TsExpressionParser; ASheetName: String);
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
    property SheetName: String read FSheetName;
  end;

  { TsCellExprNode }
  TsCellExprNode = class(TsExprNode)
  private
    FWorksheet: TsWorksheet;
    FRow, FCol: Cardinal;
    FFlags: TsRelFlags;
    FCell: PCell;
    FIsRef: Boolean;
    FOtherSheet: Boolean;
  protected
    function GetCol: Cardinal;
    function GetRow: Cardinal;
    function GetSheetIndex: Integer;
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ARow, ACol: Cardinal; AFlags: TsRelFlags; OtherSheet: Boolean); overload;
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ACellString: String; Othersheet: Boolean); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
    property Worksheet: TsWorksheet read FWorksheet;
  end;

                (*
  { TsSheetCellExprNode }
  TsSheetCellExprNode = class(TsBasicCellExprNode)
  protected
    function GetSheetIndex: Integer;
  public
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ACellString: String); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: string; override;
  end;            *)


  { TsCellRangeExprNode }

  TsCellRangeIndex = 1..2;

  TsCellRangeExprNode = class(TsExprNode)
  private
    FWorksheet: TsWorksheet;
    FRow: array[TsCellRangeIndex] of Cardinal;
    FCol: array[TsCellRangeIndex] of Cardinal;
    FFlags: TsRelFlags;
  protected
    function GetCol(AIndex: TsCellRangeIndex): Cardinal;
    function GetRow(AIndex: TsCellRangeIndex): Cardinal;
    procedure GetNodeValue(out Result: TsExpressionResult); override;
  public
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ACellRangeString: String); overload;
    constructor Create(AParser: TsExpressionParser; AWorksheet: TsWorksheet;
      ARow1,ACol1, ARow2,ACol2: Cardinal; AFlags: TsRelFlags); overload;
    function AsRPNItem(ANext: PRPNItem): PRPNItem; override;
    function AsString: String; override;
    procedure Check; override;
    function NodeType: TsResultType; override;
    property Worksheet: TsWorksheet read FWorksheet;
  end;

  { TsExpressionScanner }
  TsExpressionScanner = class(TObject)
    FSource : String;
    LSource,
    FPos: Integer;
    FChar: PChar;
    FToken: String;
    FTokenType: TsTokenType;
    FSheetNameTerminator: Char;
    FSavedSheetNameTerminator: Char;
  private
    FParser: TsExpressionParser;
    function GetCurrentChar: Char;
    procedure ScanError(Msg: String);
  protected
    procedure SetSource(const AValue: String); virtual;
    function DoError: TsTokenType;
    function DoIdentifier: TsTokenType;
    function DoNumber: TsTokenType;
    function DoDelimiter: TsTokenType;
//    function DoSquareBracket: TsTokenType;
    function DoString: TsTokenType;
    function NextPos: Char; // inline;
    procedure SkipWhiteSpace; // inline;
    function IsWordDelim(C: Char): Boolean; // inline;
    function IsDelim(C: Char): Boolean; // inline;
    function IsDigit(C: Char): Boolean; // inline;
    function IsAlpha(C: Char): Boolean; // inline;
  public
    constructor Create(AParser: TsExpressionParser);
    function GetToken: TsTokenType;
    property Token: String read FToken;
    property TokenType: TsTokenType read FTokenType;
    property Source: String read FSource write SetSource;
    property Pos: Integer read FPos;
    property CurrentChar: Char read GetCurrentChar;
    property SheetnameTerminator: char read FSheetNameTerminator write FSheetNameTerminator;
  end;

  EExprScanner = class(Exception);

  { TsExpressionParser }
  TsExpressionParser = class
  private
    FBuiltIns: TsBuiltInExprCategories;
    FExpression: String;
    FScanner: TsExpressionScanner;
    FExprNode: TsExprNode;
    FIdentifiers: TsExprIdentifierDefs;
    FHashList: TFPHashObjectlist;
    FDirty: Boolean;
    FWorksheet: TsWorksheet;
    FDialect: TsFormulaDialect;
    FSourceCell: PCell;
    FDestCell: PCell;
//    FActiveCell: PCell;
    procedure CheckEOF;
//    procedure CheckNodes(var ALeft, ARight: TsExprNode);
//    function ConvertNode(Todo: TsExprNode; ToType: TsResultType): TsExprNode;
    function GetAsBoolean: Boolean;
    function GetAsDateTime: TDateTime;
    function GetAsFloat: TsExprFloat;
    function GetAsInteger: Int64;
    function GetAsString: String;
    function GetRPNFormula: TsRPNFormula;
//    function MatchNodes(Todo, Match: TsExprNode): TsExprNode;
    procedure SetBuiltIns(const AValue: TsBuiltInExprCategories);
    procedure SetDialect(const AValue: TsFormulaDialect);
    procedure SetIdentifiers(const AValue: TsExprIdentifierDefs);
    procedure SetRPNFormula(const AFormula: TsRPNFormula);

  protected
    FFormatSettings: TFormatSettings;
    class function BuiltinExpressionManager: TsBuiltInExpressionManager;
    function BuildStringFormula(AFormatSettings: TFormatSettings): String;
    procedure ParserError(Msg: String);
    function GetExpression: String;
    function GetLocalizedExpression(const AFormatSettings: TFormatSettings): String; virtual;
    procedure SetExpression(const AValue: String);
    procedure SetLocalizedExpression(const AFormatSettings: TFormatSettings;
      const AValue: String); virtual;

    procedure CheckResultType(const Res: TsExpressionResult;
      AType: TsResultType); inline;
    function CurrentToken: String;
    function CurrentOrEOFToken: String;
    function GetToken: TsTokenType;
    function Level1: TsExprNode;
    function Level2: TsExprNode;
    function Level3: TsExprNode;
    function Level4: TsExprNode;
    function Level5: TsExprNode;
    function Level6: TsExprNode;
    function Level7: TsExprNode;
    function Primitive: TsExprNode;
    function TokenType: TsTokenType;
    procedure CreateHashList;
    property Scanner: TsExpressionScanner read FScanner;
    property ExprNode: TsExprNode read FExprNode;
    property Dirty: Boolean read FDirty;

  public
    constructor Create(AWorksheet: TsWorksheet); virtual;
    destructor Destroy; override;
    function IdentifierByName(AName: ShortString): TsExprIdentifierDef; virtual;
    procedure Clear;
    function CopyMode: Boolean;
    function Evaluate: TsExpressionResult;
    procedure EvaluateExpression(out Result: TsExpressionResult);
    procedure PrepareCopyMode(ASourceCell, ADestCell: PCell);
    function ResultType: TsResultType;

    property AsFloat: TsExprFloat read GetAsFloat;
    property AsInteger: Int64 read GetAsInteger;
    property AsString: String read GetAsString;
    property AsBoolean: Boolean read GetAsBoolean;
    property AsDateTime: TDateTime read GetAsDateTime;
    // The expression to parse
    property Expression: String read GetExpression write SetExpression;
    property LocalizedExpression[AFormatSettings: TFormatSettings]: String
        read GetLocalizedExpression write SetLocalizedExpression;
    property RPNFormula: TsRPNFormula read GetRPNFormula write SetRPNFormula;
    property Identifiers: TsExprIdentifierDefs read FIdentifiers write SetIdentifiers;
    property BuiltIns: TsBuiltInExprCategories read FBuiltIns write SetBuiltIns;
//    property ActiveCell: PCell read FActiveCell write FActiveCell;
    property Worksheet: TsWorksheet read FWorksheet;
    property Dialect: TsFormulaDialect read FDialect write SetDialect;
  end;

  TsSpreadsheetParser = class(TsExpressionParser)
  public
    constructor Create(AWorksheet: TsWorksheet); override;
  end;


  { TsBuiltInExpressionManager }
  TsBuiltInExpressionManager = class(TComponent)
  private
    FDefs: TsExprIdentifierDefs;
    function GetCount: Integer;
    function GetI(AIndex: Integer): TsBuiltInExprIdentifierDef;
  protected
    property Defs: TsExprIdentifierDefs read FDefs;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function IndexOfIdentifier(const AName: ShortString): Integer;
    function FindIdentifier(const AName: ShortString): TsBuiltInExprIdentifierDef;
    function IdentifierByExcelCode(const AExcelCode: Integer): TsBuiltInExprIdentifierDef;
    function IdentifierByName(const AName: ShortString): TsBuiltInExprIdentifierDef;
    function AddVariable(const ACategory: TsBuiltInExprCategory; const AName: ShortString;
      AResultType: TsResultType; AValue: String): TsBuiltInExprIdentifierDef;
    function AddBooleanVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: Boolean): TsBuiltInExprIdentifierDef;
    function AddIntegerVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: Integer): TsBuiltInExprIdentifierDef;
    function AddFloatVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: TsExprFloat): TsBuiltInExprIdentifierDef;
    function AddStringVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: String): TsBuiltInExprIdentifierDef;
    function AddDateTimeVariable(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; AValue: TDateTime): TsBuiltInExprIdentifierDef;
    function AddFunction(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; const AResultType: Char; const AParamTypes: String;
      const AExcelCode: Integer; ACallBack: TsExprFunctionCallBack): TsBuiltInExprIdentifierDef;
    function AddFunction(const ACategory: TsBuiltInExprCategory;
      const AName: ShortString; const AResultType: Char; const AParamTypes: String;
      const AExcelCode: Integer; ACallBack: TsExprFunctionEvent): TsBuiltInExprIdentifierDef;
    property IdentifierCount: Integer read GetCount;
    property Identifiers[AIndex: Integer]: TsBuiltInExprIdentifierDef read GetI;
  end;

  EExprParser = class(Exception);
  ECalcEngine = class(Exception);

function TokenName(AToken: TsTokenType): String;
function ResultTypeName(AResult: TsResultType): String;
function CharToResultType(C: Char): TsResultType;
function BuiltinIdentifiers: TsBuiltInExpressionManager;
function ArgToBoolean(Arg: TsExpressionResult): Boolean;
function ArgToCell(Arg: TsExpressionResult): PCell;
function ArgToDateTime(Arg: TsExpressionResult): TDateTime;
function ArgToInt(Arg: TsExpressionResult): Integer;
function ArgToFloat(Arg: TsExpressionResult): TsExprFloat;
function ArgToString(Arg: TsExpressionResult): String;
procedure ArgsToFloatArray(const Args: TsExprParameterArray; out AData: TsExprFloatArray);
function BooleanResult(AValue: Boolean): TsExpressionResult;
function CellResult(AValue: String): TsExpressionResult; overload;
function CellResult(ACellRow, ACellCol: Cardinal): TsExpressionResult; overload;
function DateTimeResult(AValue: TDateTime): TsExpressionResult;
function EmptyResult: TsExpressionResult;
function ErrorResult(const AValue: TsErrorValue): TsExpressionResult;
function FloatResult(const AValue: TsExprFloat): TsExpressionResult;
function IntegerResult(const AValue: Integer): TsExpressionResult;
function IsInteger(const AValue: TsExpressionResult): Boolean;
function IsString(const AValue: TsExpressionResult): Boolean;
function StringResult(const AValue: String): TsExpressionResult;

procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallBack: TsExprFunctionCallBack); overload;
procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallBack: TsExprFunctionEvent); overload;

var
  ExprFormatSettings: TFormatSettings;

const
  HYPERLINK_SEPARATOR = '|#@#|';  // Separats link and caption parts of a hyperlink

const
  AllBuiltIns = [bcMath, bcStatistics, bcStrings, bcLogical, bcDateTime, bcLookup,
    bcInfo, bcUser];


implementation

uses
  typinfo, math, lazutf8, dateutils, fpsutils, fpsfunc, fpsStrings;

const
  cNull = #0;
  cDoubleQuote = '"';
  cError = '#';

  Digits         = ['0'..'9'];   // + decimalseparator
  WhiteSpace     = [' ', #13, #10, #9];
  Operators      = ['+', '-', '<', '>', '=', '/', '*', '&', '%', '^'];
  Delimiters     = Operators + ['(', ')'];  // + listseparator
  Symbols        = Delimiters;
  WordDelimiters = WhiteSpace + Symbols;


{ ---------------------------------------------------------------------
  Auxiliary functions
  ---------------------------------------------------------------------}

procedure RaiseParserError(Msg: String);
begin
  raise EExprParser.Create(Msg);
end;

procedure RaiseParserError(Fmt: String; Args: Array of const);
begin
  raise EExprParser.CreateFmt(Fmt, Args);
end;

function TokenName(AToken: TsTokenType): String;
begin
  Result := GetEnumName(TypeInfo(TsTokenType), ord(AToken));
end;

function ResultTypeName(AResult: TsResultType): String;
begin
  Result := GetEnumName(TypeInfo(TsResultType), ord(AResult));
end;

function CharToResultType(C: Char): TsResultType;
begin
  case Upcase(C) of
    'S' : Result := rtString;
    'D' : Result := rtDateTime;
    'B' : Result := rtBoolean;
    'I' : Result := rtInteger;
    'F' : Result := rtFloat;
    'R' : Result := rtCellRange;
    'C' : Result := rtCell;
    '?' : Result := rtAny;
  else
    RaiseParserError(rsInvalidResultCharacter, [C]);
  end;
end;

var
  BuiltIns: TsBuiltInExpressionManager = nil;

function BuiltinIdentifiers: TsBuiltInExpressionManager;
begin
  If (BuiltIns = nil) then
    BuiltIns := TsBuiltInExpressionManager.Create(nil);
  Result := BuiltIns;
end;

procedure FreeBuiltIns;
begin
  FreeAndNil(Builtins);
end;


{------------------------------------------------------------------------------}
{  TsExpressionScanner                                                        }
{------------------------------------------------------------------------------}

constructor TsExpressionScanner.Create(AParser: TsExpressionParser);
begin
  Source := '';
  FParser := AParser;
  FSheetnameTerminator := '!';
  FSavedSheetNameTerminator := '!';
end;

function TsExpressionScanner.DoDelimiter: TsTokenType;
var
  B : Boolean;
  C, D : Char;
begin
  C := FChar^;
  FToken := C;
  B := C in ['<', '>'];
  D := C;
  C := NextPos;

  if B and (C in ['=', '>']) then
  begin
    FToken := FToken + C;
    NextPos;
    If D = '>' then
      Result := ttLargerThanEqual
    else if C = '>' then
      Result := ttNotEqual
    else
      Result := ttLessThanEqual;
  end
  else
  if D = FParser.FFormatSettings.ListSeparator then
    Result := ttListSep
  else
    case D of
      '+' : Result := ttPlus;
      '-' : Result := ttMinus;
      '*' : Result := ttMul;
      '/' : Result := ttDiv;
      '^' : Result := ttPower;
      '%' : Result := ttPercent;
      '&' : Result := ttConcat;
      '<' : Result := ttLessThan;
      '>' : Result := ttLargerThan;
      '=' : Result := ttEqual;
      '(' : Result := ttLeft;
      ')' : Result := ttRight;
  //    ',' : Result := ttComma;
    else
      ScanError(Format(rsUnknownDelimiter, [D]));
    end;
end;

function TsExpressionScanner.DoError: TsTokenType;
var
  C: Char;
begin
  C := CurrentChar;
  while (C in ['A', 'D', 'E', 'F', 'I', 'L', 'M', 'N', 'O', 'R', 'U', 'V', '0', '!', '?', '/', '#']) do
//  while (C in ['D','I','V','/','0', 'N', 'U', 'L', 'V', 'A', 'E', 'R', 'F', 'M', '!', '?']) do
//  while ((not IsWordDelim(C) or (C in ['/', '0', '!', '?'])) and (C <> cNull) do
  begin
    FToken := FToken + C;
    C := NextPos;
  end;
  Result := ttError;
end;

function TsExpressionScanner.DoIdentifier: TsTokenType;
var
  C: Char;
  S: String;
  row, row2: Cardinal;
  col, col2: Cardinal;
  flags: TsRelFlags;
begin
  C := CurrentChar;
  if C = FSheetNameTerminator then C := NextPos;
  while (not IsWordDelim(C)) and (C <> cNull) and (C <> FSheetNameTerminator) do
  begin
    if ((FParser.Dialect = fdOpenDocument) and (C = ']')) then begin
      C := NextPos;
      FSheetNameTerminator := FSavedSheetNameTerminator;
      break;
    end;
    FToken := FToken + C;
    C := NextPos;
  end;

  if C = FSheetNameTerminator then
  begin
    C := NextPos;
    result := ttSheetName;
    exit;
  end;

  S := LowerCase(FToken);
  if ParseCellString(S, row, col, flags) and (C <> '(') then
    Result := ttCell
  else if ParseCellRangeString(S, row, col, row2, col2, flags) and (C <> '(') then
    Result := ttCellRange
  else if (S = 'true') and (C <> '(') then
    Result := ttTrue
  else if (S = 'false') and (C <> '(') then
    Result := ttFalse
  else
    Result := ttIdentifier;
end;

function TsExpressionScanner.DoNumber: TsTokenType;
var
  C: Char;
  X: TsExprFloat;
  prevC: Char;
begin
  C := CurrentChar;
  prevC := #0;
  while (not IsWordDelim(C) or (prevC = 'E')) and (C <> cNull) do
  begin
    if not ( IsDigit(C)
             or ((FToken <> '') and (Upcase(C) = 'E'))
             or ((FToken <> '') and (C in ['+', '-']) and (prevC = 'E'))
           )
    then
      ScanError(Format(rsInvalidNumberChar, [C]));
    FToken := FToken+C;
    prevC := Upcase(C);
    C := NextPos;
  end;
  if not TryStrToFloat(FToken, X, FParser.FFormatSettings) then
    ScanError(Format(rsInvalidNumber, [FToken]));
  Result := ttNumber;
end;
                                    (*
{ Scans until closing square bracket is reached. In OpenDocument, this is
  a cell or cell range identifier. }
function TsExpressionScanner.DoSquareBracket: TsTokenType;
var
  C: Char;
  r1,c1,r2,c2: Cardinal;
  flags: TsRelFlags;
  isRange: Boolean;
  sheetName: String;
begin
  isRange := false;
  FToken := '';
  sheetName := '';
  C := NextPos;
  while (C <> ']') do
  begin
    case C of
      cNull: ScanError(rsUnexpectedEndOfExpression);
      '.'  : begin
               sheetName := FToken;
               FToken := '';
             end;
      ':'  : begin isRange := true; FToken := FToken + C; end;
      else   FToken := FToken + C;
    end;
    C := NextPos;
  end;
  C := NextPos;

  if sheetName <> '' then begin

  if isRange then
  begin
    if ParseCellRangeString(FToken, r1, c1, r2, c2, flags) then
      Result := ttCellRange
    else
      Result := ttError;
//      ScanError(Format(SErrInvalidCellRange, [FToken]));
  end else
  begin
    if ParseCellString(FToken, r1, c1, flags) then
      Result := ttCell
    else
      Result := ttError;
//      ScanError(Format(SErrInvalidCell, [FToken]));
  end;
end;*)


function TsExpressionScanner.DoString: TsTokenType;

  function TerminatingChar(C: Char): boolean;
  begin
    Result := (C = cNull)
          or ((C = cDoubleQuote) and
               not ((FPos < LSource) and (FSource[FPos+1] = cDoubleQuote)));
  end;

var
  C: Char;
begin
  FToken := '';
  C := NextPos;
  while not TerminatingChar(C) do
  begin
    FToken := FToken + C;
    if C = cDoubleQuote then
      NextPos;
    C := NextPos;
  end;
  if (C = cNull) then
    ScanError(rsBadQuotes);
  Result := ttString;
  FTokenType := Result;
  NextPos;
end;

function TsExpressionScanner.GetCurrentChar: Char;
begin
  if FChar <> nil then
    Result := FChar^
  else
    Result := #0;
end;

function TsExpressionScanner.GetToken: TsTokenType;
var
  C: Char;
begin
  FToken := '';
  SkipWhiteSpace;
  C := FChar^;
  if (FParser.Dialect = fdOpenDocument) and (C = '[') then begin
    FSavedSheetNameTerminator := FSheetNameTerminator;
    FSheetNameTerminator := '.';
    C := NextPos;
    Result := DoIdentifier
//    Result := DoSquareBracket
  end
  else if C = cNull then
    Result := ttEOF
  else if IsDelim(C) then
    Result := DoDelimiter
  else if (C = cDoubleQuote) then
    Result := DoString
  else if IsDigit(C) then
    Result := DoNumber
  else if (C = cError) then
    Result := DoError
  else if IsAlpha(C) or (C = '$') then
    Result := DoIdentifier
  else
    ScanError(Format(rsUnknownCharacter, [FPos, C]));
  FTokenType := Result;
end;

function TsExpressionScanner.IsAlpha(C: Char): Boolean;
begin
  Result := C in ['A'..'Z', 'a'..'z'];
end;

function TsExpressionScanner.IsDelim(C: Char): Boolean;
begin
  Result := (C in Delimiters) or (C = FParser.FFormatSettings.ListSeparator);
end;

function TsExpressionScanner.IsDigit(C: Char): Boolean;
begin
  Result := (C in Digits) or (C = FParser.FFormatSettings.DecimalSeparator);
end;

function TsExpressionScanner.IsWordDelim(C: Char): Boolean;
begin
  Result := (C in WordDelimiters) or (C = FParser.FFormatSettings.ListSeparator);
end;

function TsExpressionScanner.NextPos: Char;
begin
  Inc(FPos);
  Inc(FChar);
  Result := FChar^;
end;

procedure TsExpressionScanner.ScanError(Msg: String);
begin
  raise EExprScanner.Create(Msg)
end;

procedure TsExpressionScanner.SetSource(const AValue: String);
begin
  FSource := AValue;
  LSource := Length(FSource);
  FTokenType := ttEOF;
  if LSource = 0 then
    FPos := 0
  else
    FPos := 1;
  FChar := PChar(FSource);
  FToken := '';
end;

procedure TsExpressionScanner.SkipWhiteSpace;
begin
  while (FChar^ in WhiteSpace) and (FPos <= LSource) do
    NextPos;
end;


{------------------------------------------------------------------------------}
{  TsExpressionParser                                                         }
{------------------------------------------------------------------------------}

constructor TsExpressionParser.Create(AWorksheet: TsWorksheet);
begin
  inherited Create;
  FDialect := fdExcelA1;
  FWorksheet := AWorksheet;
  FIdentifiers := TsExprIdentifierDefs.Create(TsExprIdentifierDef);
  FIdentifiers.FParser := Self;
  FScanner := TsExpressionScanner.Create(self);
  FHashList := TFPHashObjectList.Create(False);
end;

destructor TsExpressionParser.Destroy;
begin
  FreeAndNil(FHashList);
  FreeAndNil(FExprNode);
  FreeAndNil(FIdentifiers);
  FreeAndNil(FScanner);
  inherited Destroy;
end;

{ Constructs the string formula from the tree of expression nodes. Gets the
  decimal and list separator from the formatsettings provided. }
function TsExpressionParser.BuildStringFormula(AFormatSettings: TFormatSettings): String;
begin
  ExprFormatSettings := AFormatSettings;
  if FExprNode = nil then
    Result := ''
  else
  begin
    FFormatSettings := AFormatSettings;
    Result := FExprNode.AsString;
  end;
end;

class function TsExpressionParser.BuiltinExpressionManager: TsBuiltInExpressionManager;
begin
  Result := BuiltinIdentifiers;
end;

procedure TsExpressionParser.CheckEOF;
begin
  if (TokenType = ttEOF) then
    ParserError(rsUnexpectedEndOfExpression);
end;

procedure TsExpressionParser.CheckResultType(const Res: TsExpressionResult;
  AType: TsResultType); inline;
begin
  if (Res.ResultType <> AType) then
    RaiseParserError(rsInvalidResultType, [ResultTypeName(Res.ResultType)]);
end;

procedure TsExpressionParser.Clear;
begin
  FExpression := '';
  FHashList.Clear;
  FreeAndNil(FExprNode);
end;

{ Prepares copy mode: The formula is contained in ASourceCell and will be
  modified such as seen from ADestCell. }
procedure TsExpressionParser.PrepareCopyMode(ASourceCell, ADestCell: PCell);
begin
  FSourceCell := ASourceCell;
  FDestCell := ADestCell;
end;

{ Signals that the parser is in "CopyMode", i.e. there is are source and
  destination cells. All relative references in the formula of the source cell
  habe to be adapted as seen from the destination cell. }
function TsExpressionParser.CopyMode: Boolean;
begin
  Result := (FDestCell <> nil) and (FSourceCell <> nil);
end;

procedure TsExpressionParser.CreateHashList;
var
  ID: TsExprIdentifierDef;
  BID: TsBuiltInExprIdentifierDef;
  i: Integer;
  M: TsBuiltInExpressionManager;
begin
  FHashList.Clear;
  // Builtins
  M := BuiltinExpressionManager;
  If (FBuiltins <> []) and Assigned(M) then
    for i:=0 to M.IdentifierCount-1 do
    begin
      BID := M.Identifiers[i];
      If BID.Category in FBuiltins then
        FHashList.Add(UpperCase(BID.Name), BID);
    end;
  // User
  for i:=0 to FIdentifiers.Count-1 do
  begin
    ID := FIdentifiers[i];
    FHashList.Add(UpperCase(ID.Name), ID);
  end;
  FDirty := False;
end;

function TsExpressionParser.CurrentToken: String;
begin
  Result := FScanner.Token;
end;

function TsExpressionParser.CurrentOrEOFToken: String;
begin
  if (FScanner.TokenType = ttEOF) or (FScanner.Token = '') then
    Result := 'end of formula'
  else
    Result := FScanner.Token;
end;

function TsExpressionParser.Evaluate: TsExpressionResult;
begin
  EvaluateExpression(Result);
end;

procedure TsExpressionParser.EvaluateExpression(out Result: TsExpressionResult);
begin
  if (FExpression = '') then
    ParserError(rsExpressionEmpty);
  if not Assigned(FExprNode) then
    ParserError(rsErrorInExpression);
  FExprNode.GetNodeValue(Result);
end;

function TsExpressionParser.GetAsBoolean: Boolean;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtBoolean);
  Result := Res.ResBoolean;
end;

function TsExpressionParser.GetAsDateTime: TDateTime;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtDateTime);
  Result := Res.ResDatetime;
end;

function TsExpressionParser.GetAsFloat: TsExprFloat;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtFloat);
  Result := Res.ResFloat;
end;

function TsExpressionParser.GetAsInteger: Int64;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtInteger);
  Result := Res.ResInteger;
end;

function TsExpressionParser.GetAsString: String;
var
  Res: TsExpressionResult;
begin
  EvaluateExpression(Res);
  CheckResultType(Res, rtString);
  Result := Res.ResString;
end;

function TsExpressionParser.GetRPNFormula: TsRPNFormula;
begin
  Result := CreateRPNFormula(FExprNode.AsRPNItem(nil), true);
end;

function TsExpressionParser.GetToken: TsTokenType;
begin
  Result := FScanner.GetToken;
end;

function TsExpressionParser.IdentifierByName(AName: ShortString): TsExprIdentifierDef;
begin
  if FDirty then
    CreateHashList;
  Result := TsExprIdentifierDef(FHashList.Find(UpperCase(AName)));
end;

function TsExpressionParser.Level1: TsExprNode;
{
var
  tt: TsTokenType;
  Right: TsExprNode;
  }
begin
{$ifdef debugexpr}Writeln('Level 1 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
{
  if TokenType = ttNot then
  begin
    GetToken;
    CheckEOF;
    Right := Level2;
    Result := TsNotExprNode.Create(Right);
  end
  else
  }
  Result := Level2;
{
  try

    while (TokenType in [ttAnd, ttOr, ttXor]) do
    begin
      tt := TokenType;
      GetToken;
      CheckEOF;
      Right := Level2;
      case tt of
        ttOr  : Result := TsBinaryOrExprNode.Create(Result, Right);
        ttAnd : Result := TsBinaryAndExprNode.Create(Result, Right);
        ttXor : Result := TsBinaryXorExprNode.Create(Result, Right);
      else
        ParserError(SErrUnknownBooleanOp)
      end;
    end;
  except
    Result.Free;
    raise;
  end;
}
end;

function TsExpressionParser.Level2: TsExprNode;
var
  right: TsExprNode;
  tt: TsTokenType;
  C: TsBinaryOperationExprNodeClass;
begin
{$ifdef debugexpr} Writeln('Level 2 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  Result := Level3;
  try
    if (TokenType in ttComparisons) then
    begin
      tt := TokenType;
      GetToken;
      CheckEOF;
      Right := Level3;
      case tt of
        ttLessthan         : C := TsLessExprNode;
        ttLessthanEqual    : C := TsLessEqualExprNode;
        ttLargerThan       : C := TsGreaterExprNode;
        ttLargerThanEqual  : C := TsGreaterEqualExprNode;
        ttEqual            : C := TsEqualExprNode;
        ttNotEqual         : C := TsNotEqualExprNode;
      else
        ParserError(rsUnknownComparison)
      end;
      Result := C.Create(self, Result, right);
    end;
  except
    Result.Free;
    raise;
  end;
end;

function TsExpressionParser.Level3: TsExprNode;
var
  tt: TsTokenType;
  right: TsExprNode;
begin
{$ifdef debugexpr} Writeln('Level 3 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  Result := Level4;
  try
    while TokenType in [ttPlus, ttMinus, ttConcat] do begin
      tt := TokenType;
      GetToken;
      CheckEOF;
      right := Level4;
      case tt of
        ttPlus  : Result := TsAddExprNode.Create(self, Result, right);
        ttMinus : Result := TsSubtractExprNode.Create(self, Result, right);
        ttConcat: Result := TsConcatExprNode.Create(self, Result, right);
      end;
    end;
  except
    Result.Free;
    raise;
  end;
end;

function TsExpressionParser.Level4: TsExprNode;
var
  tt: TsTokenType;
  right: TsExprNode;
begin
{$ifdef debugexpr} Writeln('Level 4 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  Result := Level5;
  try
    while (TokenType in [ttMul, ttDiv]) do
    begin
      tt := TokenType;
      GetToken;
      right := Level5;
      case tt of
        ttMul : Result := TsMultiplyExprNode.Create(self, Result, right);
        ttDiv : Result := TsDivideExprNode.Create(self, Result, right);
      end;
    end;
  except
    Result.Free;
    Raise;
  end;
end;

function TsExpressionParser.Level5: TsExprNode;
var
  right: TsExprNode;
begin
  {$ifdef debugexpr} Writeln('Level 5 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  Result := Level6;
  try
    while (TokenType = ttPower) do
    begin
      GetToken;
      right := Level6;
      Result := TsPowerExprNode.Create(self, Result, right);
    end;
  except
    Result.Free;
    Raise;
  end;
end;

function TsExpressionParser.Level6: TsExprNode;
var
  signs: String;
  i: Integer;
begin
{$ifdef debugexpr} Writeln('Level 6 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  signs := '';
  while (TokenType in [ttPlus, ttMinus]) do
  begin
    case TokenType of
      ttPlus  : signs := signs + '+';
      ttMinus : signs := signs + '-';
    end;
    GetToken;
  end;
  Result := Level7;
  i := Length(signs);
  while (i > 0) do begin
    case signs[i] of
      '+':  Result := TsUPlusExprNode.Create(self, Result);
      '-':  Result := TsUMinusExprNode.Create(self, Result);
    end;
    dec(i);
  end;

  while TokenType = ttPercent do begin
    Result := TsPercentExprNode.Create(self, Result);
    GetToken;
  end;
end;

function TsExpressionParser.Level7: TsExprNode;
var
  currToken: String;
begin
{$ifdef debugexpr} Writeln('Level 7 ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  if (TokenType = ttLeft) then
  begin
    GetToken;
    Result := TsParenthesisExprNode.Create(self, Level1);
    try
      if (TokenType <> ttRight) then begin
        currToken := CurrentToken;
        if TokenType = ttEOF then currToken := 'end of formula';
        ParserError(Format(rsRightBracketExpected, [SCanner.Pos, currToken]));
      end;
      GetToken;
    except
      Result.Free;
      raise;
    end;
  end
  else
    Result := Primitive;
end;

procedure TsExpressionParser.ParserError(Msg: String);
begin
  raise EExprParser.Create(Msg);
end;

function TsExpressionParser.Primitive: TsExprNode;
var
  I: Int64;
  X: TsExprFloat;
  lCount: Integer;
  ID: TsExprIdentifierDef;
  Args: TsExprArgumentArray;
  AI: Integer;
  optional: Boolean;
  token: String;
  prevTokenType: TsTokenType;
  sheetname: String;
  sheet: TsWorksheet;
begin
{$ifdef debugexpr} Writeln('Primitive : ',TokenName(TokenType),': ',CurrentToken);{$endif debugexpr}
  SetLength(Args, 0);
  if (TokenType = ttNumber) then
  begin
    if TryStrToInt64(CurrentToken, I) then
      Result := TsConstExprNode.CreateInteger(self, I)
    else
    if TryStrToFloat(CurrentToken, X, FFormatSettings) then
      Result := TsConstExprNode.CreateFloat(self, X)
    else
      ParserError(Format(rsInvalidFloat, [CurrentToken]));
  end
  else if (TokenType = ttTrue) then
    Result := TsConstExprNode.CreateBoolean(self, true)
  else if (TokenType = ttFalse) then
    Result := TsConstExprNode.CreateBoolean(self, false)
  else if (TokenType = ttString) then
    Result := TsConstExprNode.CreateString(self, CurrentToken)
  else if (TokenType = ttCell) then
    Result := TsCellExprNode.Create(self, FWorksheet, CurrentToken, false)
  else if (TokenType = ttSheetName) then begin
    sheetName := CurrentToken;
    GetToken;
    if TokenType = ttCell then begin
      sheet := FWorksheet.Workbook.GetWorksheetByName(sheetName);
      if sheet = nil then
        sheet := FWorksheet.Workbook.AddWorksheet(sheetName, true);
      Result := TsCellExprNode.Create(self, sheet, CurrentToken, true)
    end;
  end
  else if (TokenType = ttCellRange) then
    Result := TsCellRangeExprNode.Create(self, FWorksheet, CurrentToken)
  else if (TokenType = ttError) then
    Result := TsConstExprNode.CreateError(self, CurrentToken)
  else if not (TokenType in [ttIdentifier]) then
    ParserError(Format(rsUnknownTokenAtPos, [Scanner.Pos, CurrentToken]))
  else
  begin
    token := Uppercase(CurrentToken);
    ID := self.IdentifierByName(token);
    if (ID = nil) then
      ParserError(Format(rsUnknownIdentifier, [token]));
    if (ID.IdentifierType in [itFunctionCallBack, itFunctionHandler]) then
    begin
      lCount := ID.ArgumentCount;
      if lCount = 0 then  // we have to handle the () here, it will be skipped below.
      begin
        GetToken;
        if (TokenType <> ttLeft) then
          ParserError(Format(rsLeftBracketExpected, [Scanner.Pos, CurrentOrEOFToken]));
        GetToken;
        if (TokenType <> ttRight) then
          ParserError(Format(rsRightBracketExpected, [Scanner.Pos, CurrentOrEOFToken]));
        SetLength(Args, 0);
      end;
    end
    else
      lCount := 0;

    // Parse arguments.
    // Negative is for variable number of arguments, where Abs(value) is the minimum number of arguments
    if (lCount <> 0) then
    begin
      GetToken;
      if (TokenType <> ttLeft) then
        ParserError(Format(rsLeftBracketExpected, [Scanner.Pos, CurrentOrEofToken]));
      SetLength(Args, abs(lCount));
      AI := 0;
      try
        repeat
          prevTokenType := TokenType;
          GetToken;
          // Check if we must enlarge the argument array
          if (lCount < 0) and (AI = Length(Args)) then
          begin
            SetLength(Args, AI+1);
            Args[AI] := nil;
          end;
          if (prevTokenType in [ttLeft, ttListSep]) and (TokenType in [ttListSep, ttRight]) then
          begin
            Args[AI] := TsMissingArgExprNode.Create;
            inc(AI);
            Continue;
          end;
          Args[AI] := Level1;
          inc(AI);
          optional := ID.IsOptionalArgument(AI+1);
          if not optional then
          begin
            if (TokenType <> ttListSep) then
              if (AI < abs(lCount)) then
                ParserError(Format(rsCommaExpected, [Scanner.Pos, CurrentOrEofToken]))
          end;
        until (AI = lCount) or (((lCount < 0) or optional) and (TokenType = ttRight));
        if TokenType <> ttRight then
          ParserError(Format(rsRightBracketExpected, [Scanner.Pos, CurrentOrEofToken]));
        if AI < abs(lCount) then
          SetLength(Args, AI);
      except
        on E: Exception do
        begin
          dec(AI);
          while (AI >= 0) do
          begin
            FreeAndNil(Args[Ai]);
            dec(AI);
          end;
          raise;
        end;
      end;
    end;
    case ID.IdentifierType of
      itVariable:
        Result := TsVariableExprNode.CreateIdentifier(self, ID);
      itFunctionCallBack:
        Result := TsFunctionCallBackExprNode.CreateFunction(self, ID, Args);
      itFunctionHandler:
        Result := TFPFunctionEventHandlerExprNode.CreateFunction(self, ID, Args);
    end;
  end;
  GetToken;
end;

function TsExpressionParser.ResultType: TsResultType;
begin
  if not Assigned(FExprNode) then
    ParserError(rsErrorInExpression);
  Result := FExprNode.NodeType;;
end;

procedure TsExpressionParser.SetBuiltIns(const AValue: TsBuiltInExprCategories);
begin
  if FBuiltIns = AValue then
    exit;
  FBuiltIns := AValue;
  FDirty := true;
end;

function TsExpressionParser.GetExpression: String;
var
  fs: TFormatsettings;
begin
  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';
  fs.ListSeparator := ',';
  Result := BuildStringFormula(fs);
end;

function TsExpressionParser.GetLocalizedExpression(const AFormatSettings: TFormatSettings): String;
begin
  ExprFormatSettings := AFormatSettings;
  Result := BuildStringFormula(AFormatSettings);
end;

procedure TsExpressionParser.SetDialect(const AValue: TsFormulaDialect);
begin
  if FDialect = AValue then exit;
  FDialect := AValue;
  {
  if FScanner <> nil then
    case FDialect of
      fdExcelA1, fdExcelR1C1: FScanner.SheetNameTerminator := '!';
      fdOpenDocument: FScanner.Sheetnameterminator := '.';
      else raise Exception.Create('TsExpressionParser.SetDialect: Dialect not supported.');
    end;
    }
end;

procedure TsExpressionParser.SetExpression(const AValue: String);
var
  fs: TFormatSettings;
begin
  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';
  fs.ListSeparator := ',';
  SetLocalizedExpression(fs, AValue);
end;

procedure TsExpressionParser.SetLocalizedExpression(const AFormatSettings: TFormatSettings;
  const AValue: String);
begin
  if FExpression = AValue then
    exit;
  FFormatSettings := AFormatSettings;
  ExprFormatSettings := AFormatSettings;
  FExpression := AValue;
  if (AValue <> '') and (AValue[1] = '=') then
    FScanner.Source := Copy(AValue, 2, Length(AValue))
  else
    FScanner.Source := AValue;
  FreeAndNil(FExprNode);
  if (FExpression <> '') then
  begin
    GetToken;
    FExprNode := Level1;
    if (TokenType <> ttEOF) then
      ParserError(Format(rsUnterminatedExpression, [Scanner.Pos, CurrentToken]));
    FExprNode.Check;
  end;
end;

procedure TsExpressionParser.SetIdentifiers(const AValue: TsExprIdentifierDefs);
begin
  FIdentifiers.Assign(AValue)
end;

procedure TsExpressionParser.SetRPNFormula(const AFormula: TsRPNFormula);

  procedure CreateNodeFromRPN(var ANode: TsExprNode; var AIndex: Integer);
  var
    left: TsExprNode = nil;
    right: TsExprNode = nil;
    operand: TsExprNode = nil;
    fek: TFEKind;
    r,c, r2,c2: Cardinal;
    idx: Integer;
    flags: TsRelFlags;
    ID: TsExprIdentifierDef;
    i, n: Integer;
    args: TsExprArgumentArray;
    sheet: TsWorksheet;
  begin
    if AIndex < 0 then
      exit;

    fek := AFormula[AIndex].ElementKind;

    case fek of
      fekCell, fekCellRef:
        begin
          r := AFormula[AIndex].Row;
          c := AFormula[AIndex].Col;
          if (LongInt(r) < 0) or (LongInt(c) < 0) then
            ANode := TsConstExprNode.CreateError(self, errIllegalRef)
          else
          begin
            flags := AFormula[AIndex].RelFlags;
            ANode := TsCellExprNode.Create(self, FWorksheet, r, c, flags, false);
          end;
          dec(AIndex);
        end;
      fekCell3D:
        begin
          idx := AFormula[AIndex].Sheet;
          r := AFormula[AIndex].Row;
          c := AFormula[AIndex].Col;
          if (LongInt(r) < 0) or (LongInt(c) < 0) then
            ANode := TsConstExprNode.CreateError(self, errIllegalRef)
          else
          begin
            flags := AFormula[AIndex].RelFlags;
            sheet := FWorksheet.Workbook.GetWorksheetByIndex(idx);
            ANode := TsCellExprNode.Create(Self, sheet, r, c, flags, true);
          end;
          dec(AIndex);
        end;
      fekCellRange:
        begin
          r := AFormula[AIndex].Row;
          c := AFormula[AIndex].Col;
          r2 := AFormula[AIndex].Row2;
          c2 := AFormula[AIndex].Col2;
          flags := AFormula[AIndex].RelFlags;
          ANode := TsCellRangeExprNode.Create(self, FWorksheet, r, c, r2, c2, flags);
          dec(AIndex);
        end;
      fekNum:
        begin
          ANode := TsConstExprNode.CreateFloat(self, AFormula[AIndex].DoubleValue);
          dec(AIndex);
        end;
      fekInteger:
        begin
          ANode := TsConstExprNode.CreateInteger(self, AFormula[AIndex].IntValue);
          dec(AIndex);
        end;
      fekString:
        begin
          ANode := TsConstExprNode.CreateString(self, AFormula[AIndex].StringValue);
          dec(AIndex);
        end;
      fekBool:
        begin
          ANode := TsConstExprNode.CreateBoolean(self, AFormula[AIndex].DoubleValue <> 0.0);
          dec(AIndex);
        end;
      fekErr:
        begin
          ANode := TsConstExprNode.CreateError(self, TsErrorValue(AFormula[AIndex].IntValue));
          dec(AIndex);
        end;
      fekMissingArg:
        begin
          ANode := TsMissingArgExprNode.Create;
          dec(AIndex);
        end;

      // unary operations
      fekPercent, fekUMinus, fekUPlus, fekParen:
        begin
          dec(AIndex);
          CreateNodeFromRPN(operand, AIndex);
          case fek of
            fekPercent : ANode := TsPercentExprNode.Create(self, operand);
            fekUMinus  : ANode := TsUMinusExprNode.Create(self, operand);
            fekUPlus   : ANode := TsUPlusExprNode.Create(self, operand);
            fekParen   : ANode := TsParenthesisExprNode.Create(self, operand);
          end;
        end;

      // binary operations
      fekAdd, fekSub, fekMul, fekDiv,
      fekPower, fekConcat,
      fekEqual, fekNotEqual,
      fekGreater, fekGreaterEqual,
      fekLess, fekLessEqual:
        begin
          dec(AIndex);
          CreateNodeFromRPN(right, AIndex);
          CreateNodeFromRPN(left, AIndex);
          //CheckNodes(left, right);
          case fek of
            fekAdd         : ANode := TsAddExprNode.Create(self, left, right);
            fekSub         : ANode := TsSubtractExprNode.Create(self, left, right);
            fekMul         : ANode := TsMultiplyExprNode.Create(self, left, right);
            fekDiv         : ANode := TsDivideExprNode.Create(self, left, right);
            fekPower       : ANode := TsPowerExprNode.Create(self, left, right);
            fekConcat      : ANode := tsConcatExprNode.Create(self, left, right);
            fekEqual       : ANode := TsEqualExprNode.Create(self, left, right);
            fekNotEqual    : ANode := TsNotEqualExprNode.Create(self, left, right);
            fekGreater     : ANode := TsGreaterExprNode.Create(self, left, right);
            fekGreaterEqual: ANode := TsGreaterEqualExprNode.Create(self, left, right);
            fekLess        : ANode := TsLessExprNode.Create(self, left, right);
            fekLessEqual   : ANode := tsLessEqualExprNode.Create(self, left, right);
          end;
        end;

      // functions
      fekFunc:
        begin
          ID := self.IdentifierByName(AFormula[AIndex].FuncName);
          if ID = nil then
          begin
            ParserError(Format(rsUnknownIdentifier, [AFormula[AIndex].FuncName]));
            dec(AIndex);
          end else
          begin
            if ID.HasFixedArgumentCount then
              n := ID.ArgumentCount
            else
              n := AFormula[AIndex].ParamsNum;
            dec(AIndex);
            SetLength(args, n);
            for i:=n-1 downto 0 do
              CreateNodeFromRPN(args[i], AIndex);
            case ID.IdentifierType of
              itVariable:
                ANode := TsVariableExprNode.CreateIdentifier(self, ID);
              itFunctionCallBack:
                ANode := TsFunctionCallBackExprNode.CreateFunction(self, ID, args);
              itFunctionHandler:
                ANode := TFPFunctionEventHandlerExprNode.CreateFunction(self, ID, args);
            end;
          end;
        end;

    end;  //case
  end; //begin

var
  index: Integer;
begin
  FExpression := '';
  FreeAndNil(FExprNode);
  index := Length(AFormula)-1;
  CreateNodeFromRPN(FExprNode, index);
  if Assigned(FExprNode) then FExprNode.Check;
end;

function TsExpressionParser.TokenType: TsTokenType;
begin
  Result := FScanner.TokenType;
end;


{------------------------------------------------------------------------------}
{  TsSpreadsheetParser                                                         }
{------------------------------------------------------------------------------}

constructor TsSpreadsheetParser.Create(AWorksheet: TsWorksheet);
begin
  inherited Create(AWorksheet);
  BuiltIns := AllBuiltIns;
end;


{------------------------------------------------------------------------------}
{  TsExprIdentifierDefs                                                        }
{------------------------------------------------------------------------------}

function TsExprIdentifierDefs.AddBooleanVariable(const AName: ShortString;
  AValue: Boolean): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtBoolean;
  Result.FValue.ResBoolean := AValue;
end;

function TsExprIdentifierDefs.AddDateTimeVariable(const AName: ShortString;
  AValue: TDateTime): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtDateTime;
  Result.FValue.ResDateTime := AValue;
end;

function TsExprIdentifierDefs.AddFloatVariable(const AName: ShortString;
  AValue: TsExprFloat): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtFloat;
  Result.FValue.ResFloat := AValue;
end;

function TsExprIdentifierDefs.AddFunction(const AName: ShortString;
  const AResultType: Char; const AParamTypes: String; const AExcelCode: Integer;
  ACallBack: TsExprFunctionCallBack): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.Name := AName;
  Result.IdentifierType := itFunctionCallBack;
  Result.ResultType := CharToResultType(AResultType);
  Result.ExcelCode := AExcelCode;
  Result.FOnGetValueCB := ACallBack;
  if (Length(AParamTypes) > 0) and (AParamTypes[Length(AParamTypes)]='+') then
  begin
    Result.ParameterTypes := Copy(AParamTypes, 1, Length(AParamTypes)-1);
    Result.VariableArgumentCount := true;
  end else
    Result.ParameterTypes := AParamTypes;
end;

function TsExprIdentifierDefs.AddFunction(const AName: ShortString;
  const AResultType: Char; const AParamTypes: String; const AExcelCode: Integer;
  ACallBack: TsExprFunctionEvent): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.Name := AName;
  Result.IdentifierType := itFunctionHandler;
  Result.ResultType := CharToResultType(AResultType);
  Result.ExcelCode := AExcelCode;
  Result.FOnGetValue := ACallBack;
  if (Length(AParamTypes) > 0) and (AParamTypes[Length(AParamTypes)]='+') then
  begin
    Result.ParameterTypes := Copy(AParamTypes, 1, Length(AParamTypes)-1);
    Result.VariableArgumentCount := true;
  end else
    Result.ParameterTypes := AParamTypes;
end;

function TsExprIdentifierDefs.AddIntegerVariable(const AName: ShortString;
  AValue: Integer): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtInteger;
  Result.FValue.ResInteger := AValue;
end;

function TsExprIdentifierDefs.AddStringVariable(const AName: ShortString;
  AValue: String): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := rtString;
  Result.FValue.ResString := AValue;
end;

function TsExprIdentifierDefs.AddVariable(const AName: ShortString;
  AResultType: TsResultType; AValue: String): TsExprIdentifierDef;
begin
  Result := Add as TsExprIdentifierDef;
  Result.IdentifierType := itVariable;
  Result.Name := AName;
  Result.ResultType := AResultType;
  Result.Value := AValue;
end;

function TsExprIdentifierDefs.FindIdentifier(const AName: ShortString
  ): TsExprIdentifierDef;
var
  I: Integer;
begin
  I := IndexOfIdentifier(AName);
  if (I = -1) then
    Result := nil
  else
    Result := GetI(I);
end;

function TsExprIdentifierDefs.GetI(AIndex : Integer): TsExprIdentifierDef;
begin
  Result := TsExprIdentifierDef(Items[AIndex]);
end;

function TsExprIdentifierDefs.IdentifierByExcelCode(const AExcelCode: Integer
  ): TsExprIdentifierDef;
var
  I: Integer;
begin
  I := IndexOfIdentifier(AExcelCode);
  if I = -1 then
    Result := nil
  else
    Result := GetI(I);
end;

function TsExprIdentifierDefs.IdentifierByName(const AName: ShortString
  ): TsExprIdentifierDef;
begin
  Result := FindIdentifier(AName);
  if (Result = nil) then
    RaiseParserError(rsUnknownIdentifier, [AName]);
end;

function TsExprIdentifierDefs.IndexOfIdentifier(const AName: ShortString): Integer;
begin
  Result := Count-1;
  while (Result >= 0) and (CompareText(GetI(Result).Name, AName) <> 0) do
    dec(Result);
end;

function TsExprIdentifierDefs.IndexOfIdentifier(const AExcelCode: Integer): Integer;
var
  ID: TsExprIdentifierDef;
begin
  Result := Count-1;
  while (Result >= 0) do begin
    ID := GetI(Result);
    if ID.ExcelCode = AExcelCode then exit;
    dec(Result);
  end;
end;

procedure TsExprIdentifierDefs.SetI(AIndex: Integer;
  const AValue: TsExprIdentifierDef);
begin
  Items[AIndex] := AValue;
end;

procedure TsExprIdentifierDefs.Update(Item: TCollectionItem);
begin
  Unused(Item);
  if Assigned(FParser) then
    FParser.FDirty := true;
end;


{------------------------------------------------------------------------------}
{  TsExprIdentifierDef                                                        }
{------------------------------------------------------------------------------}

function TsExprIdentifierDef.ArgumentCount: Integer;
begin
  if FVariableArgumentCount then
    Result := -Length(FArgumentTypes)
  else
    Result := Length(FArgumentTypes);
end;

procedure TsExprIdentifierDef.Assign(Source: TPersistent);
var
  EID: TsExprIdentifierDef;
begin
  if (Source is TsExprIdentifierDef) then
  begin
    EID := Source as TsExprIdentifierDef;
    FStringValue := EID.FStringValue;
    FValue := EID.FValue;
    FArgumentTypes := EID.FArgumentTypes;
    FVariableArgumentCount := EID.FVariableArgumentCount;
    FExcelCode := EID.ExcelCode;
    FIDType := EID.FIDType;
    FName := EID.FName;
    FOnGetValue := EID.FOnGetValue;
    FOnGetValueCB := EID.FOnGetValueCB;
  end
  else
    inherited Assign(Source);
end;

procedure TsExprIdentifierDef.CheckResultType(const AType: TsResultType);
begin
  if (FValue.ResultType <> AType) then
    RaiseParserError(rsInvalidResultType, [ResultTypeName(AType)])
end;

procedure TsExprIdentifierDef.CheckVariable;
begin
  if Identifiertype <> itVariable then
    RaiseParserError(rsNoVariable, [Name]);
end;

function TsExprIdentifierDef.GetAsBoolean: Boolean;
begin
  CheckResultType(rtBoolean);
  CheckVariable;
  Result := FValue.ResBoolean;
end;

function TsExprIdentifierDef.GetAsDateTime: TDateTime;
begin
  CheckResultType(rtDateTime);
  CheckVariable;
  Result := FValue.ResDateTime;
end;

function TsExprIdentifierDef.GetAsFloat: TsExprFloat;
begin
  CheckResultType(rtFloat);
  CheckVariable;
  Result := FValue.ResFloat;
end;

function TsExprIdentifierDef.GetAsInteger: Int64;
begin
  CheckResultType(rtInteger);
  CheckVariable;
  Result := FValue.ResInteger;
end;

function TsExprIdentifierDef.GetAsString: String;
begin
  CheckResultType(rtString);
  CheckVariable;
  Result := FValue.ResString;
end;

function TsExprIdentifierDef.GetFormatSettings: TFormatSettings;
begin
  Result := TsExprIdentifierDefs(Collection).Parser.FFormatSettings;
end;

function TsExprIdentifierDef.GetResultType: TsResultType;
begin
  Result := FValue.ResultType;
end;

function TsExprIdentifierDef.GetValue: String;
begin
  case FValue.ResultType of
    rtBoolean  : if FValue.ResBoolean then
                   Result := 'TRUE'
                 else
                   Result := 'FALSE';
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtFloat    : Result := FloatToStr(FValue.ResFloat, GetFormatSettings);
    rtDateTime : Result := FormatDateTime('cccc', FValue.ResDateTime, GetFormatSettings);
    rtString   : Result := FValue.ResString;
  end;
end;

{ Returns true if the epxression has a fixed number of arguments. }
function TsExprIdentifierDef.HasFixedArgumentCount: Boolean;
var
  i: Integer;
begin
  if FVariableArgumentCount then
    Result := false
  else
  begin
    for i:= 1 to Length(FArgumentTypes) do
      if IsOptionalArgument(i) then
      begin
        Result := false;
        exit;
      end;
    Result := true;
  end;
end;

{ Checks whether an argument is optional. Index number starts at 1.
  Optional arguments are lower-case characters in the argument list. }
function TsExprIdentifierDef.IsOptionalArgument(AIndex: Integer): Boolean;
begin
  Result := (AIndex <= Length(FArgumentTypes))
    and (UpCase(FArgumentTypes[AIndex]) <> FArgumentTypes[AIndex]);
end;

procedure TsExprIdentifierDef.SetArgumentTypes(const AValue: String);
var
  i: integer;
begin
  if FArgumentTypes = AValue then
    exit;
  for i:=1 to Length(AValue) do
    CharToResultType(AValue[i]);
  FArgumentTypes := AValue;
end;

procedure TsExprIdentifierDef.SetAsBoolean(const AValue: Boolean);
begin
  CheckVariable;
  CheckResultType(rtBoolean);
  FValue.ResBoolean := AValue;
end;

procedure TsExprIdentifierDef.SetAsDateTime(const AValue: TDateTime);
begin
  CheckVariable;
  CheckResultType(rtDateTime);
  FValue.ResDateTime := AValue;
end;

procedure TsExprIdentifierDef.SetAsFloat(const AValue: TsExprFloat);
begin
  CheckVariable;
  CheckResultType(rtFloat);
  FValue.ResFloat := AValue;
end;

procedure TsExprIdentifierDef.SetAsInteger(const AValue: Int64);
begin
  CheckVariable;
  CheckResultType(rtInteger);
  FValue.ResInteger := AValue;
end;

procedure TsExprIdentifierDef.SetAsString(const AValue: String);
begin
  CheckVariable;
  CheckResultType(rtString);
  FValue.ResString := AValue;
end;

procedure TsExprIdentifierDef.SetName(const AValue: ShortString);
begin
  if FName = AValue then
    exit;
  if (AValue <> '') then
    if Assigned(Collection) and (TsExprIdentifierDefs(Collection).IndexOfIdentifier(AValue) <> -1) then
      RaiseParserError(rsDuplicateIdentifier,[AValue]);
  FName := AValue;
end;

procedure TsExprIdentifierDef.SetResultType(const AValue: TsResultType);
begin
  if AValue <> FValue.ResultType then
  begin
    FValue.ResultType := AValue;
    SetValue(FStringValue);
  end;
end;

procedure TsExprIdentifierDef.SetValue(const AValue: String);
begin
  FStringValue := AValue;
  if (AValue <> '') then
    case FValue.ResultType of
      rtBoolean  : FValue.ResBoolean := (FStringValue='True');
      rtInteger  : FValue.ResInteger := StrToInt(AValue);
      rtFloat    : FValue.ResFloat := StrToFloat(AValue, GetFormatSettings);
      rtDateTime : FValue.ResDateTime := StrToDateTime(AValue, GetFormatSettings);
      rtString   : FValue.ResString := AValue;
    end
  else
    case FValue.ResultType of
      rtBoolean  : FValue.ResBoolean := false;
      rtInteger  : FValue.ResInteger := 0;
      rtFloat    : FValue.ResFloat := 0.0;
      rtDateTime : FValue.ResDateTime := 0;
      rtString   : FValue.ResString := '';
    end
end;


{------------------------------------------------------------------------------}
{  TsBuiltInExpressionManager                                                         }
{------------------------------------------------------------------------------}

constructor TsBuiltInExpressionManager.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDefs := TsExprIdentifierDefs.Create(TsBuiltInExprIdentifierDef)
end;

destructor TsBuiltInExpressionManager.Destroy;
begin
  FreeAndNil(FDefs);
  inherited Destroy;
end;

function TsBuiltInExpressionManager.AddVariable(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; AResultType: TsResultType; AValue: String
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.Addvariable(AName, AResultType, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddBooleanVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: Boolean
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddBooleanvariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddDateTimeVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: TDateTime
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddDateTimeVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFloatVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString;
  AValue: TsExprFloat): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFloatVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFunction(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; const AResultType: Char; const AParamTypes: String;
  const AExcelCode: Integer; ACallBack: TsExprFunctionCallBack): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFunction(AName, AResultType,
    AParamTypes, AExcelCode, ACallBack));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddFunction(const ACategory: TsBuiltInExprCategory;
  const AName: ShortString; const AResultType: Char; const AParamTypes: String;
  const AExcelCode: Integer; ACallBack: TsExprFunctionEvent): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddFunction(AName, AResultType,
    AParamTypes, AExcelCode, ACallBack));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddIntegerVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: Integer
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddIntegerVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.AddStringVariable(
  const ACategory: TsBuiltInExprCategory; const AName: ShortString; AValue: String
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.AddStringVariable(AName, AValue));
  Result.Category := ACategory;
end;

function TsBuiltInExpressionManager.FindIdentifier(const AName: ShortString
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.FindIdentifier(AName));
end;

function TsBuiltInExpressionManager.GetCount: Integer;
begin
  Result := FDefs.Count;
end;

function TsBuiltInExpressionManager.GetI(AIndex: Integer): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs[Aindex])
end;

function TsBuiltInExpressionManager.IdentifierByExcelCode(const AExcelCode: Integer
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.IdentifierByExcelCode(AExcelCode));
end;

function TsBuiltInExpressionManager.IdentifierByName(const AName: ShortString
  ): TsBuiltInExprIdentifierDef;
begin
  Result := TsBuiltInExprIdentifierDef(FDefs.IdentifierByName(AName));
end;

function TsBuiltInExpressionManager.IndexOfIdentifier(const AName: ShortString): Integer;
begin
  Result := FDefs.IndexOfIdentifier(AName);
end;


{------------------------------------------------------------------------------}
{  Various Nodes                                                               }
{------------------------------------------------------------------------------}

{ TsExprNode }

procedure TsExprNode.Check;
begin
end;

function TsExprNode.HasError(out AResult: TsExpressionResult): Boolean;
begin
  GetNodeValue(AResult);
  if AResult.ResultType = rtError then
  begin
    Result := true;
    AResult := ErrorResult(AResult.ResError);
  end else
    Result := false;
end;

function TsExprNode.NodeValue: TsExpressionResult;
begin
  GetNodeValue(Result);
end;


{ TsUnaryOperationExprNode }

constructor TsUnaryOperationExprNode.Create(AParser: TsExpressionParser;
  AOperand: TsExprNode);
begin
  FParser := AParser;
  FOperand := AOperand;
end;

destructor TsUnaryOperationExprNode.Destroy;
begin
  FreeAndNil(FOperand);
  inherited Destroy;
end;

procedure TsUnaryOperationExprNode.Check;
begin
  if not Assigned(Operand) then
    RaiseParserError(rsNoOperand, [Self.ClassName]);
end;


{ TsBinaryOperationExprNode }

constructor TsBinaryOperationExprNode.Create(AParser: TsExpressionParser;
  ALeft, ARight: TsExprNode);
begin
  FParser := AParser;
  FLeft := ALeft;
  FRight := ARight;
end;

destructor TsBinaryOperationExprNode.Destroy;
begin
  FreeAndNil(FLeft);
  FreeAndNil(FRight);
  inherited Destroy;
end;

function TsBinaryOperationExprNode.HasError(out AResult: TsExpressionResult): Boolean;
begin
  Result := Left.HasError(AResult) or Right.HasError(AResult);
end;


{ TsBooleanOperationExprNode }

function TsBooleanOperationExprNode.NodeType: TsResultType;
begin
  Result := Left.NodeType;
end;


{ TsConstExprNode }

constructor TsConstExprNode.CreateString(AParser: TsExpressionParser;
  AValue: String);
begin
  FParser := AParser;
  FValue.ResultType := rtString;
  FValue.ResString := AValue;
end;

constructor TsConstExprNode.CreateInteger(AParser: TsExpressionParser;
  AValue: Int64);
begin
  FParser := AParser;
  FValue.ResultType := rtInteger;
  FValue.ResInteger := AValue;
end;

constructor TsConstExprNode.CreateDateTime(AParser: TsExpressionParser;
  AValue: TDateTime);
begin
  FParser := AParser;
  FValue.ResultType := rtDateTime;
  FValue.ResDateTime := AValue;
end;

constructor TsConstExprNode.CreateFloat(AParser: TsExpressionParser;
  AValue: TsExprFloat);
begin
  FParser := AParser;
  FValue.ResultType := rtFloat;
  FValue.ResFloat := AValue;
end;

constructor TsConstExprNode.CreateBoolean(AParser: TsExpressionParser;
  AValue: Boolean);
begin
  FParser := AParser;
  FValue.ResultType := rtBoolean;
  FValue.ResBoolean := AValue;
end;

constructor TsConstExprNode.CreateError(AParser: TsExpressionParser;
  AValue: TsErrorValue);
begin
  FParser := AParser;
  FValue.ResultType := rtError;
  FValue.ResError := AValue;
end;

constructor TsConstExprNode.CreateError(AParser: TsExpressionParser;
  AValue: String);
var
  err: TsErrorValue;
begin
  // Don't check for equal strings. If, for example, the column A of a cell
  // reference A1 is deleted Excel replaces the A by '#REF!', i.e the
  // reference becomes '#REF!1' (with the 1 at the end)!
  if pos('#NULL!', AValue) > 0 then
//  if AValue = '#NULL!' then
    err := errEmptyIntersection
  else if Pos('#DIV/0!', AValue) > 0 then
//  else if AValue = '#DIV/0!' then
    err := errDivideByZero
//  else if AValue = '#VALUE!' then
  else if Pos('#VALUE!', AValue) > 0 then
    err := errWrongType
//  else if AValue = '#REF!' then
  else if Pos('#REF!', AValue) > 0 then
    err := errIllegalRef
//  else if AValue = '#NAME?' then
  else if Pos('#NAME?', AValue) > 0 then
    err := errWrongName
//  else if AValue = '#NUM!' then
  else if Pos('#NUM!', AValue) > 0 then
    err := errOverflow
//  else if AValue = '#N/A' then
  else if Pos('#N/A', AValue) > 0 then
    err := errArgError
//  else if AValue = '#FORMULA?' then
  else if Pos('#FORMULA?', AValue) > 0 then
    err := errFormulaNotSupported
  else
    AParser.ParserError('Unknown error type.');
  CreateError(AParser, err);
end;

function TsConstExprNode.NodeType: TsResultType;
begin
  Result := FValue.ResultType;
end;

procedure TsConstExprNode.GetNodeValue(out AResult: TsExpressionResult);
begin
  AResult := FValue;
end;

function TsConstExprNode.AsString: string;
begin
  case NodeType of
    rtString   : Result := cDoubleQuote + FValue.ResString + cDoubleQuote;
    rtInteger  : Result := IntToStr(FValue.ResInteger);
    rtDateTime : Result := '''' + FormatDateTime('cccc', FValue.ResDateTime, Parser.FFormatSettings) + '''';    // Probably wrong !!!
    rtBoolean  : if FValue.ResBoolean then Result := 'TRUE' else Result := 'FALSE';
    rtFloat    : Result := FloatToStr(FValue.ResFloat, Parser.FFormatSettings);
    rtError    : Result := GetErrorValueStr(FValue.ResError);
  end;
end;

function TsConstExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  case NodeType of
    rtString   : Result := RPNString(FValue.ResString, ANext);
    rtInteger  : Result := RPNInteger(FValue.ResInteger, ANext);
    rtDateTime : Result := RPNNumber(FValue.ResDateTime, ANext);
    rtBoolean  : Result := RPNBool(FValue.ResBoolean, ANext);
    rtFloat    : Result := RPNNumber(FValue.ResFloat, ANext);
    rtError    : Result := RPNErr(FValue.ResError, ANext);
  end;
end;


{ TsMissingExprNode }

function TsMissingArgExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNMissingARg(ANext);
end;

function TsMissingArgExprNode.AsString: String;
begin
  Result := '';
end;

procedure TsMissingArgExprNode.GetNodeValue(out AResult: TsExpressionResult);
begin
  AResult.ResultType := rtMissingArg;
  AResult.ResInteger := 0;
end;

function TsMissingArgExprNode.NodeType: TsResultType;
begin
  Result := rtMissingArg;
end;

{ TsUPlusExprNode }

function TsUPlusExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekUPlus,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsUPlusExprNode.AsString: String;
begin
  Result := '+' + TrimLeft(Operand.AsString);
end;

procedure TsUPlusExprNode.GetNodeValue(out Result: TsExpressionResult);
var
  cell: PCell;
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtInteger, rtFloat, rtError:
      exit;
    rtCell:
      begin
        cell := ArgToCell(Result);
        if cell = nil then
          Result := FloatResult(0.0)
        else
        if cell^.ContentType = cctNumber then
        begin
          if frac(cell^.NumberValue) = 0.0 then
            Result := IntegerResult(trunc(cell^.NumberValue))
          else
            Result := FloatResult(cell^.NumberValue);
        end;
      end;
    rtEmpty:
      Result := FloatResult(0.0);
    else
      Result := ErrorResult(errWrongType);
  end;
end;

function TsUPlusExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;


{ TsUMinusExprNode }

function TsUMinusExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekUMinus,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsUMinusExprNode.AsString: String;
begin
  Result := '-' + TrimLeft(Operand.AsString);
end;

procedure TsUMinusExprNode.GetNodeValue(out Result: TsExpressionResult);
var
  cell: PCell;
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtError:
      exit;
    rtFloat:
      Result := FloatResult(-Result.ResFloat);
    rtInteger:
      Result := IntegerResult(-Result.ResInteger);
    rtCell:
      begin
        cell := ArgToCell(Result);
        if (cell <> nil) and (cell^.ContentType = cctNumber) then
        begin
          if frac(cell^.NumberValue) = 0.0 then
            Result := IntegerResult(-trunc(cell^.NumberValue))
          else
            Result := FloatResult(cell^.NumberValue);
        end else
          Result := FloatResult(0.0);
      end;
    rtEmpty:
      Result := FloatResult(0.0);
    else
      Result := ErrorResult(errWrongType);
  end;
end;

function TsUMinusExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;


{ TsPercentExprNode }

function TsPercentExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekPercent,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsPercentExprNode.AsString: String;
begin
  Result := Operand.AsString + '%';
end;

procedure TsPercentExprNode.Check;
const
  AllowedTokens = [rtInteger, rtFloat, rtCell, rtEmpty, rtError];
begin
  inherited;
  if not (Operand.NodeType in AllowedTokens) then
    RaiseParserError(rsNoPercentOperation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsPercentExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Operand.GetNodeValue(Result);
  case Result.ResultType of
    rtError:
      exit;
    rtFloat, rtInteger, rtCell:
      Result := FloatResult(ArgToFloat(Result)*0.01);
    else
      Result := ErrorResult(errWrongType);
  end;
end;

function TsPercentExprNode.NodeType: TsResultType;
begin
  Result := rtFloat;
end;


{ TsParenthesisExprNode }

function TsParenthesisExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekParen,
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsParenthesisExprNode.AsString: String;
begin
  Result := '(' + Operand.AsString + ')';
end;

function TsParenthesisExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;

procedure TsParenthesisExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Result := Operand.NodeValue;
end;

         (*
{ TsNotExprNode }

function TsNotExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc('NOT',
    Operand.AsRPNItem(
    ANext
  ));
end;

function TsNotExprNode.AsString: String;
begin
  Result := 'not ' + Operand.AsString;
end;

procedure TsNotExprNode.Check;
const
  AllowedTokens = [rtBoolean, rtEmpty, rtError];
begin
  if not (Operand.NodeType in AllowedTokens) then
    RaiseParserError(SErrNoNotOperation, [ResultTypeName(Operand.NodeType), Operand.AsString])
end;

procedure TsNotExprNode.GetNodeValue(out AResult: TsExpressionResult);
begin
  Operand.GetNodeValue(AResult);
  case AResult.ResultType of
    rtBoolean : AResult.ResBoolean := not AResult.ResBoolean;
    rtEmpty   : AResult := BooleanResult(true);    // This is according to Excel
  end
end;

function TsNotExprNode.NodeType: TsResultType;
begin
  Result := Operand.NodeType;
end;
           *)

{ TsBooleanResultExprNode }

function TsBooleanResultExprNode.NodeType: TsResultType;
begin
  Result := rtBoolean;
end;


{ TsEqualExprNode }

function TsEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '=' + Right.AsString;
end;

procedure TsEqualExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  if Left.HasError(AResult) and Right.HasError(AResult) then
  begin
    AResult := BooleanResult(LRes.ResError = RRes.ResError);
    exit;
  end;

  if HasError(AResult) then
    exit;

  if IsString(LRes) and IsString(RRes) then
    AResult := BooleanResult(ArgToString(LRes) = ArgToString(RRes))
  else
    AResult := BooleanResult(ArgToFloat(LRes) = ArgToFloat(RRes));
end;


{ TsNotEqualExprNode }

function TsNotEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekNotEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsNotEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '<>' + Right.AsString;
end;

procedure TsNotEqualExprNode.GetNodeValue(out AResult: TsExpressionResult);
begin
  inherited GetNodeValue(AResult);
  if AResult.ResultType <> rtError then
    AResult.ResBoolean := not AResult.ResBoolean;
end;


{ TsLessExprNode }

function TsLessExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekLess,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsLessExprNode.AsString: string;
begin
  Result := Left.AsString + '<' + Right.AsString;
end;

procedure TsLessExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  if IsString(LRes) and IsString(RRes) then
    AResult := BooleanResult(ArgToString(LRes) < ArgToString(RRes))
  else
    AResult := BooleanResult(ArgToFloat(LRes) < ArgToFloat(RRes));
end;


{ TsGreaterExprNode }

function TsGreaterExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekGreater,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsGreaterExprNode.AsString: string;
begin
  Result := Left.AsString + '>' + Right.AsString;
end;

procedure TsGreaterExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  if IsString(LRes) and IsString(RRes) then
    AResult := BooleanResult(ArgToString(LRes) > ArgToString(RRes))
  else
    AResult := BooleanResult(ArgToFloat(LRes) > ArgToFloat(RRes));
end;


{ TsGreaterEqualExprNode }

function TsGreaterEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekGreaterEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsGreaterEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '>=' + Right.AsString;
end;

procedure TsGreaterEqualExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  if IsString(LRes) and IsString(RRes) then
    AResult := BooleanResult(ArgToString(LRes) >= ArgToString(RRes))
  else
    AResult := BooleanResult(ArgToFloat(LRes) >= ArgToFloat(RRes));
end;


{ TsLessEqualExprNode }

function TsLessEqualExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekLessEqual,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsLessEqualExprNode.AsString: string;
begin
  Result := Left.AsString + '<=' + Right.AsString;
end;

procedure TsLessEqualExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  if IsString(LRes) and IsString(RRes) then
    AResult := BooleanResult(ArgToString(LRes) <= ArgToString(RRes))
  else
    AResult := BooleanResult(ArgToFloat(LRes) <= ArgToFloat(RRes));
end;


{ TsConcatExprNode }

function TsConcatExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekConcat,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext)));
end;

function TsConcatExprNode.AsString: string;
begin
  Result := Left.AsString + '&' + Right.AsString;
end;

procedure TsConcatExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes : TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  AResult := StringResult(ArgToString(LRes) + ArgToString(RRes));
end;

function TsConcatExprNode.NodeType: TsResultType;
begin
  Result := rtString;
end;


{ TsMathOperationExprNode }

function TsMathOperationExprNode.NodeType: TsResultType;
begin
  Result := Left.NodeType;
end;


{ TsAddExprNode }

function TsAddExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekAdd,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsAddExprNode.AsString: string;
begin
  Result := Left.AsString + '+' + Right.AsString;
end;

procedure TsAddExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  if IsInteger(LRes) and IsInteger(RRes) then
    AResult := IntegerResult(ArgToInt(LRes) + ArgToInt(RRes))
  else
    AResult := FloatResult(ArgToFloat(LRes) + ArgToFloat(RRes));
end;


{ TsSubtractExprNode }

function TsSubtractExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekSub,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsSubtractExprNode.AsString: string;
begin
  Result := Left.AsString + '-' + Right.asString;
end;

procedure TsSubtractExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  lRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  if IsInteger(LRes) and IsInteger(RRes)
  then
    AResult := IntegerResult(ArgToInt(LRes) - ArgToInt(RRes))
  else
    AResult := FloatResult(ArgToFloat(LRes) - ArgToFloat(RRes))
end;


{ TsMultiplyExprNode }

function TsMultiplyExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekMul,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsMultiplyExprNode.AsString: string;
begin
  Result := Left.AsString + '*' + Right.AsString;
end;

procedure TsMultiplyExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);
  try
    AResult := FloatResult(ArgToFloat(LRes) * ArgToFloat(RRes));
  except
    on EInvalidArgument do AResult := ErrorResult(errOverflow);
  end;
end;


{ TsDivideExprNode }

function TsDivideExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekDiv,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsDivideExprNode.AsString: string;
begin
  Result := Left.AsString + '/' + Right.asString;
end;

procedure TsDivideExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
  y: TsExprFloat;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);

  y := ArgToFloat(RRes);
  if y = 0.0 then
    AResult := ErrorResult(errDivideByZero)
  else
    try
      AResult := FloatResult(ArgToFloat(LRes) / y);
    except
      on EInvalidArgument do AResult := ErrorResult(errOverflow);
    end;
end;

function TsDivideExprNode.NodeType: TsResultType;
begin
  Result := rtFLoat;
end;


{ TsPowerExprNode }

function TsPowerExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNFunc(fekPower,
    Right.AsRPNItem(
    Left.AsRPNItem(
    ANext
  )));
end;

function TsPowerExprNode.AsString: string;
begin
  Result := Left.AsString + '^' + Right.AsString;
end;

procedure TsPowerExprNode.GetNodeValue(out AResult: TsExpressionResult);
var
  LRes, RRes: TsExpressionResult;
begin
  if HasError(AResult) then
    exit;

  Left.GetNodeValue(LRes);
  Right.GetNodeValue(RRes);
  try
    AResult := FloatResult(Power(ArgToFloat(LRes), ArgToFloat(RRes)));
  except
    on E: EInvalidArgument do AResult := ErrorResult(errOverflow);
  end;
end;

function TsPowerExprNode.NodeType: TsResultType;
begin
  Result := rtFloat;
end;


{ TsIdentifierExprNode }

constructor TsIdentifierExprNode.CreateIdentifier(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef);
begin
  FParser := AParser;
  FID := AID;
  PResult := @FID.FValue;
  FResultType := FID.ResultType;
end;

function TsIdentifierExprNode.NodeType: TsResultType;
begin
  Result := FResultType;
end;

procedure TsIdentifierExprNode.GetNodeValue(out AResult: TsExpressionResult);
begin
  AResult := PResult^;
  AResult.ResultType := FResultType;
end;


{ TsVariableExprNode }

function TsVariableExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := ANext;  // Just a dummy assignment to silence the compiler...
  RaiseParserError('Cannot handle variables for RPN, so far.');
end;

function TsVariableExprNode.AsString: string;
begin
  Result := FID.Name;
end;


{ TsFunctionExprNode }

constructor TsFunctionExprNode.CreateFunction(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef; const Args: TsExprArgumentArray);
begin
  inherited CreateIdentifier(AParser, AID);
  FArgumentNodes := Args;
  SetLength(FArgumentParams, Length(Args));
end;

destructor TsFunctionExprNode.Destroy;
var
  i: Integer;
begin
  for i:=0 to Length(FArgumentNodes)-1 do
    FreeAndNil(FArgumentNodes[i]);
  inherited Destroy;
end;

function TsFunctionExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
var
  i, n: Integer;
begin
  if FID.HasFixedArgumentCount then
    n := FID.ArgumentCount
  else
    n := Length(FArgumentNodes);
  Result := ANext;
  for i:=0 to High(FArgumentNodes) do
    Result := FArgumentNodes[i].AsRPNItem(Result);
  Result := RPNFunc(FID.Name, n, Result);
end;

function TsFunctionExprNode.AsString: String;
var
  S : String;
  i : Integer;
begin
  S := '';
  for i := 0 to Length(FArgumentNodes)-1 do
  begin
    if (S <> '') then
      S := S + Parser.FFormatSettings.ListSeparator;
    if Assigned(FArgumentNodes[i]) then
      S := S + FArgumentNodes[i].AsString;
  end;
  S := '(' + S + ')';
  Result := FID.Name + S;
end;

procedure TsFunctionExprNode.CalcParams;
var
  i : Integer;
begin
  for i := 0 to Length(FArgumentParams)-1 do
    FArgumentNodes[i].GetNodeValue(FArgumentParams[i]);
end;

procedure TsFunctionExprNode.Check;
var
  i: Integer;
  rta,                  // Parameter types passed to the function
  rtp: TsResultType;    // Parameter types expected from the parameter symbol
  lastrtp: TsResultType;
begin
  if Length(FArgumentNodes) <> FID.ArgumentCount then
  begin
    for i:=Length(FArgumentNodes)+1 to FID.ArgumentCount do
      if not FID.IsOptionalArgument(i) then
        RaiseParserError(rsInvalidArgumentCount, [FID.Name]);
  end;

  for i := 0 to Length(FArgumentNodes)-1 do
  begin
    if FArgumentNodes[i] = nil then
      Continue;

    rta := FArgumentNodes[i].NodeType;

    if i+1 <= Length(FID.ParameterTypes) then
    begin
      rtp := CharToResultType(FID.ParameterTypes[i+1]);
      lastrtp := rtp;
    end else
      rtp := lastrtp;

    if rtp = rtAny then
      Continue;
    // A "cell" can return any type --> no type conversion required here.

    if rta = rtCell then
      Continue;
                    (*
    if (rtp <> rta) and not (rta in [rtCellRange, rtError, rtEmpty]) then
    begin
      // Automatically convert integers to floats in functions that return a float
      if (rta = rtInteger) and (rtp = rtFloat) then
      begin
        FArgumentNodes[i] := TsIntToFloatExprNode(FArgumentNodes[i]);
        exit;
      end;
      // Floats are truncated automatically to integers - that's what Excel does.
      if (rta = rtFloat) and (rtp = rtInteger) then
        exit;
      RaiseParserError(SErrInvalidArgumentType, [i+1, ResultTypeName(rtp), ResultTypeName(rta)])
    end;
    *)
  end;
end;


{ TsFunctionCallBackExprNode }

constructor TsFunctionCallBackExprNode.CreateFunction(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef; const Args: TsExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValueCallBack;
end;

procedure TsFunctionCallBackExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Result.ResultType := NodeType;     // was at end!
  if Length(FArgumentParams) > 0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
end;


{ TFPFunctionEventHandlerExprNode }

constructor TFPFunctionEventHandlerExprNode.CreateFunction(AParser: TsExpressionParser;
  AID: TsExprIdentifierDef; const Args: TsExprArgumentArray);
begin
  inherited;
  FCallBack := AID.OnGetFunctionValue;
end;

procedure TFPFunctionEventHandlerExprNode.GetNodeValue(out Result: TsExpressionResult);
begin
  Result.ResultType := NodeType;    // was at end
  if Length(FArgumentParams) > 0 then
    CalcParams;
  FCallBack(Result, FArgumentParams);
end;


{ TsSheetNameExprNode }
constructor TsSheetNameExprNode.Create(AParser: TsExpressionParser;
  ASheetName: string);
begin
  FParser := AParser;
  FSheetName := ASheetName;
end;

function TsSheetNameExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := ANext;
end;

function TsSheetnameExprNode.AsString: string;
begin
  Result := '';
end;


{ TsCellExprNode }

constructor TsCellExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ARow,ACol: Cardinal; AFlags: TsRelFlags;
  OtherSheet: Boolean);
begin
  FParser := AParser;
  FWorksheet := AWorksheet;
  FRow := ARow;
  FCol := ACol;
  FFlags := AFlags;
  FCell := AWorksheet.FindCell(FRow, FCol);
  FOtherSheet := OtherSheet;
end;

constructor TsCellExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ACellString: String; OtherSheet: Boolean);
var
  r, c: Cardinal;
  flags: TsRelFlags;
begin
  ParseCellString(ACellString, r, c, flags);
  Create(AParser, AWorksheet, r, c, flags, OtherSheet);
end;

function TsCellExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  if FIsRef then
  begin
    if FOtherSheet then
      Result := RPNCellRef3D(GetSheetIndex, GetRow, GetCol, FFlags, ANext)
    else
      Result := RPNCellRef(GetRow, GetCol, FFlags, ANext)
  end else
  begin
    if FOtherSheet then
      Result := RPNCellValue3D(GetSheetIndex, GetRow, GetCol, FFlags, ANext)
    else
      Result := RPNCellValue(GetRow, GetCol, FFlags, ANext);
  end;
end;

function TsCellExprNode.AsString: string;
var
  r, c: Cardinal;
begin
  r := Getrow;
  c := GetCol;
  if FOtherSheet then
    case FParser.Dialect of
      fdExcelA1:
        Result := Format('%s!%s', [FWorksheet.Name, GetCellString(r, c, FFlags)]);
      fdExcelR1C1:
        Result := Format('%s!%s', [FWorksheet.Name,
          GetCellString_R1C1(r, c, FFlags, FParser.FSourceCell^.Row, FParser.FSourceCell^.Col)]);
      fdOpenDocument:
        Result := Format('[%s.%s]', [FWorksheet.Name, GetCellString(r, c, FFlags)]);
    end
  else
    case FParser.Dialect of
      fdExcelA1:
        Result := GetCellString(GetRow, GetCol, FFlags);
      fdExcelR1C1:
        Result := GetCellString_R1C1(GetRow, GetCol, FFlags, FParser.FSourceCell^.Row, FParser.FSourceCell^.Col);
      fdOpenDocument:
        Result := '[.' + GetCellString(GetRow, GetCol, FFlags) + ']';
    end;
end;

procedure TsCellExprNode.Check;
begin
  // Nothing to check;
end;

{ Calculates the column address of the node's cell for various cases:
  (1) Copy mode:
      The "DestCell" of the parser is the cell for which the formula is
      calculated. The "SourceCell" contains the formula. If the formula contains
      a relative address in the cell node the function calculates the row
      address of the cell represented by the node as seen from the DestCell.
      If the formula contains an absolute address the function returns the row
      address of the SourceCell.
  (2) Normal mode:
      Returns the "true" row address of the cell assigned to the formula node. }
function TsCellExprNode.GetCol: Cardinal;
begin
  Result := FCol;
  if FParser.CopyMode and (rfRelCol in FFlags) then
    Result := FCol - FParser.FSourceCell^.Col + FParser.FDestCell^.Col;
end;

procedure TsCellExprNode.GetNodeValue(out Result: TsExpressionResult);
var
  cell: PCell;
begin
  if Parser.CopyMode then
    cell := FWorksheet.FindCell(GetRow, GetCol)
  else
    cell := FCell;

  if (cell <> nil) and HasFormula(cell) then
    case FWorksheet.GetCalcState(cell) of
      csNotCalculated:
        FWorksheet.CalcFormula(cell);
      csCalculating:
        raise ECalcEngine.CreateFmt(rsCircularReference, [GetCellString(cell^.Row, cell^.Col)]);
    end;

  Result.ResultType := rtCell;
  Result.ResRow := GetRow;
  Result.ResCol := GetCol;
  Result.Worksheet := FWorksheet;
end;

{ See: GetCol }
function TsCellExprNode.GetRow: Cardinal;
begin
  Result := FRow;
  if Parser.CopyMode and (rfRelRow in FFlags) then
    Result := FRow - FParser.FSourceCell^.Row + FParser.FDestCell^.Row;
end;

function TsCellExprNode.GetSheetIndex: Integer;
var
  book: TsWorkbook;
begin
  book := FWorksheet.Workbook;
  Result := book.GetWorksheetIndex(FWorksheet);
end;

function TsCellExprNode.NodeType: TsResultType;
begin
  Result := rtCell;
end;

               (*
{ TsSheetCellExprNode }

constructor TsSheetCellExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ACellString: String);
var
  r, c: Cardinal;
  flags: TsRelFlags;
  p: Integer;
  sheetname: String;
begin
  (*
  case AParser.Dialect of
    fdExcelA1, fdExcelR1C1: p := pos('!', ACellString);
    fdOpendocument: p := pos('.', ACellString);
    else raise Exception.Create('TsSheetCellExprNode: Parser dialect not supported.');
  end;
  sheetname := copy(ACellString, 1, p-1);
  ACellString := copy(ACellString, p+1, MaxInt);
  *)
  ParseCellString(ACellString, r, c, flags);
  Create(AParser, AWorksheet, r, c, flags);
end;

function TsSheetCellExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  if FIsRef then
    Result := RPNCellRef3D(GetSheetIndex, GetRow, GetCol, FFlags, ANext)
  else
    Result := RPNCellValue3D(GetSheetIndex, GetRow, GetCol, FFlags, ANext);
end;

function TsSheetCellExprNode.AsString: String;
begin
  case FParser.Dialect of
    fdExcelA1:
      Result := Format('%s!%s', [
        FWorksheet.Name,
        GetCellString(GetRow, GetCol, FFlags)
      ]);
    fdExcelR1C1:
      Result := Format('%s!%s', [
        FWorksheet.Name,
        GetCellString_R1C1(GetRow, GetCol, FFlags, FParser.FSourceCell^.Row, FParser.FSourceCell^.Col)
      ]);
    fdOpenDocument:
      Result := Format('[%s.%s]', [
        FWorksheet.Name,
        GetCellString(GetRow, GetCol, FFlags)
      ]);
    else
      raise Exception.Create('TsSheetCellExprNode: Parser dialect not supported.');
  end;
end;

function TsSheetCellExprNode.GetSheetIndex: Integer;
var
  book: TsWorkbook;
begin
  book := FWorksheet.Workbook;
  Result := book.GetWorksheetIndex(FWorksheet);
end;
            *)

{ TsCellExprNode }


{ TsCellRangeExprNode }

constructor TsCellRangeExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ACellRangeString: String);
var
  r1, c1, r2, c2: Cardinal;
  flags: TsRelFlags;
begin
  if pos(':', ACellRangeString) = 0 then
  begin
    ParseCellString(ACellRangeString, r1, c1, flags);
    if rfRelRow in flags then Include(flags, rfRelRow2);
    if rfRelCol in flags then Include(flags, rfRelCol2);
    Create(AParser, AWorksheet, r1, c1, r1, c1, flags);
  end else
  begin
    ParseCellRangeString(ACellRangeString, r1, c1, r2, c2, flags);
    Create(AParser, AWorksheet, r1, c1, r2, c2, flags);
  end;
end;

constructor TsCellRangeExprNode.Create(AParser: TsExpressionParser;
  AWorksheet: TsWorksheet; ARow1,ACol1,ARow2,ACol2: Cardinal; AFlags: TsRelFlags);
begin
  FParser := AParser;
  FWorksheet := AWorksheet;
  FRow[1] := ARow1;
  FCol[1] := ACol1;
  FRow[2] := ARow2;
  FCol[2] := ACol2;
  FFlags := AFlags;
end;

function TsCellRangeExprNode.AsRPNItem(ANext: PRPNItem): PRPNItem;
begin
  Result := RPNCellRange(GetRow(1), GetCol(1), GetRow(2), GetCol(2), FFlags, ANext);
end;

function TsCellRangeExprNode.AsString: string;
var
  r, c: Array[TsCellRangeIndex] of Cardinal;
  i: TsCellRangeIndex;
begin
  for i in TsCellRangeIndex do
  begin
    r[i] := GetRow(i);
    c[i] := GetCol(i);
  end;
  if (r[1] = r[2]) and (c[1] = c[2]) then
    Result := GetCellString(r[1], r[1], FFlags)
  else
    Result := GetCellRangeString(r[1], c[1], r[2], c[2], FFlags);
end;

procedure TsCellRangeExprNode.Check;
begin
  // Nothing to check;
end;

{ Calculates the column address of the node's cell for various cases:
  (1) Copy mode:
      The "DestCell" of the parser is the cell for which the formula is
      calculated. The "SourceCell" contains the formula. If the formula contains
      a relative address in the cell node the function calculates the row
      address of the cell represented by the node as seen from the DestCell.
      If the formula contains an absolute address the function returns the row
      address of the SourceCell.
  (2) Normal mode:
      Returns the "true" row address of the cell assigned to the formula node. }
function TsCellRangeExprNode.GetCol(AIndex: TsCellRangeIndex): Cardinal;
begin
  Result := FCol[AIndex];
  if FParser.CopyMode and (rfRelCol in FFlags) then
    Result := FCol[AIndex] - FParser.FSourceCell^.Col + FParser.FDestCell^.Col;
end;

procedure TsCellRangeExprNode.GetNodeValue(out Result: TsExpressionResult);
var
  r, c: Array[TsCellRangeIndex] of Cardinal;
  rr, cc: Cardinal;
  i: TsCellRangeIndex;
  cell: PCell;
begin
  for i in TsCellRangeIndex do
  begin
    r[i] := GetRow(i);
    c[i] := GetCol(i);
  end;
  for rr := r[1] to r[2] do
    for cc := c[1] to c[2] do
    begin
      cell := FWorksheet.FindCell(rr, cc);
      if HasFormula(cell) then
        case FWorksheet.GetCalcState(cell) of
          csNotCalculated:
            FWorksheet.CalcFormula(cell);
          csCalculating:
            raise ECalcEngine.Create(rsCircularReference);
        end;
    end;

  Result.ResultType := rtCellRange;
  Result.ResCellRange.Row1 := r[1];
  Result.ResCellRange.Col1 := c[1];
  Result.ResCellRange.Row2 := r[2];
  Result.ResCellRange.Col2 := c[2];
  Result.Worksheet := FWorksheet;
end;

function TsCellRangeExprNode.GetRow(AIndex: TsCellRangeIndex): Cardinal;
begin
  Result := FRow[AIndex];
  if FParser.CopyMode and (rfRelRow in FFlags) then
    Result := FRow[AIndex] - FParser.FSourceCell^.Row + FParser.FDestCell^.Row;
end;

function TsCellRangeExprNode.NodeType: TsResultType;
begin
  Result := rtCellRange;
end;


{------------------------------------------------------------------------------}
{   Conversion of arguments to simple data types                               }
{------------------------------------------------------------------------------}

function ArgToBoolean(Arg: TsExpressionResult): Boolean;
var
  cell: PCell;
begin
  Result := false;
  if Arg.ResultType = rtBoolean then
    Result := Arg.ResBoolean
  else
  if (Arg.ResultType = rtCell) then begin
    cell := ArgToCell(Arg);
    if (cell <> nil) and (cell^.ContentType = cctBool) then
      Result := cell^.BoolValue;
  end;
end;

function ArgToCell(Arg: TsExpressionResult): PCell;
begin
  if Arg.ResultType = rtCell then
    Result := Arg.Worksheet.FindCell(Arg.ResRow, Arg.ResCol)
  else
    Result := nil;
end;

function ArgToInt(Arg: TsExpressionResult): Integer;
var
  cell: PCell;
begin
  Result := 0;
  case Arg.ResultType of
    rtInteger   : result := Arg.ResInteger;
    rtFloat     : result := trunc(Arg.ResFloat);
    rtDateTime  : result := trunc(Arg.ResDateTime);
    rtBoolean   : if Arg.ResBoolean then Result := 1 else Result := 0;
    rtString,
    rtHyperlink : TryStrToInt(ArgToString(Arg), Result);
    rtCell      : begin
                    cell := ArgToCell(Arg);
                    if Assigned(cell) then
                      case cell^.ContentType of
                        cctNumber    : result := trunc(cell^.NumberValue);
                        cctDateTime  : result := trunc(cell^.DateTimeValue);
                        cctBool      : if cell^.BoolValue then result := 1;
                        cctUTF8String: if not TryStrToInt(cell^.UTF8StringValue, result)
                                         then Result := 0;
                      end;
                  end;
  end;
end;

{ Utility function for the built-in math functions. Accepts also integers and
 other data types in place of floating point arguments. To be called in
 builtins or user-defined callbacks having float results or arguments. }
function ArgToFloat(Arg: TsExpressionResult): TsExprFloat;
var
  cell: PCell;
  s: String;
  fs: TFormatSettings;
begin
  Result := 0.0;
  case Arg.ResultType of
    rtInteger   : result := Arg.ResInteger;
    rtDateTime  : result := Arg.ResDateTime;
    rtFloat     : result := Arg.ResFloat;
    rtBoolean   : if Arg.ResBoolean then Result := 1.0;
    rtString,
    rtHyperlink : TryStrToFloat(ArgToString(Arg), Result);
    rtCell      : begin
                    cell := ArgToCell(Arg);
                    if Assigned(cell) then
                      case cell^.ContentType of
                        cctNumber    : Result := cell^.NumberValue;
                        cctDateTime  : Result := cell^.DateTimeValue;
                        cctBool      : if cell^.BoolValue then result := 1.0;
                        cctUTF8String: begin
                                         fs := Arg.Worksheet.Workbook.FormatSettings;
                                         s := cell^.UTF8StringValue;
                                         TryStrToFloat(s, result, fs);
                                       end;
                       end;
                   end;
  end;
end;

function ArgToDateTime(Arg: TsExpressionResult): TDateTime;
var
  cell: PCell;
  fs: TFormatSettings;
begin
  Result := 0.0;
  case Arg.ResultType of
    rtDateTime  : result := Arg.ResDateTime;
    rtInteger   : Result := Arg.ResInteger;
    rtFloat     : Result := Arg.ResFloat;
    rtBoolean   : if Arg.ResBoolean then Result := 1.0;
    rtString,
    rtHyperlink : begin
                    fs := Arg.Worksheet.Workbook.FormatSettings;
                    TryStrToDateTime(ArgToString(Arg), Result, fs);
                  end;
    rtCell      : begin
                    cell := ArgToCell(Arg);
                    if Assigned(cell) then
                      if (cell^.ContentType = cctDateTime) then
                        Result := cell^.DateTimeValue;
                  end;
  end;
end;

function ArgToString(Arg: TsExpressionResult): String;
// The Office applications are very fuzzy about data types...
var
  cell: PCell;
  fs: TFormatSettings;
  dt: TDateTime;
  p: Integer;
  s: String;
begin
  Result := '';
  case Arg.ResultType of
    rtString    : result := Arg.ResString;
    rtHyperlink : begin
                    s := Arg.ResString;
                    p := pos(HYPERLINK_SEPARATOR, s);
                    if p = 0 then
                      Result := s
                    else
                      Result := Copy(s, p + Length(HYPERLINK_SEPARATOR), Length(s));
                  end;
    rtInteger   : Result := IntToStr(Arg.ResInteger);
    rtFloat     : Result := FloatToStr(Arg.ResFloat);
    rtBoolean   : if Arg.ResBoolean then Result := '1' else Result := '0';
    rtCell      : begin
                    cell := ArgToCell(Arg);
                    if Assigned(cell) then
                      case cell^.ContentType of
                        cctUTF8String : Result := cell^.UTF8Stringvalue;
                        cctNumber     : Result := Format('%g', [cell^.NumberValue]);
                        cctBool       : if cell^.BoolValue then Result := '1' else Result := '0';
                        cctDateTime   : begin
                                          fs := Arg.Worksheet.Workbook.FormatSettings;
                                          dt := cell^.DateTimeValue;
                                          if frac(dt) = 0.0 then
                                            Result := FormatDateTime(fs.LongTimeFormat, dt, fs)
                                          else
                                          if trunc(dt) = 0 then
                                            Result := FormatDateTime(fs.ShortDateFormat, dt, fs)
                                          else
                                            Result := FormatDateTime('cc', dt, fs);
                                        end;
                      end;
                  end;
  end;
end;

procedure ArgsToFloatArray(const Args: TsExprParameterArray; out AData: TsExprFloatArray);
const
  BLOCKSIZE = 128;
var
  i, n: Integer;
  r, c: Cardinal;
  cell: PCell;
  arg: TsExpressionResult;
begin
  SetLength(AData, BLOCKSIZE);
  n := 0;
  for i:=0 to High(Args) do
  begin
    arg := Args[i];
    if arg.ResultType = rtCellRange then
      for r := arg.ResCellRange.Row1 to arg.ResCellRange.Row2 do
        for c := arg.ResCellRange.Col1 to arg.ResCellRange.Col2 do
        begin
          cell := arg.Worksheet.FindCell(r, c);
          if (cell <> nil) and (cell^.ContentType in [cctNumber, cctDateTime]) then
          begin
            case cell^.ContentType of
              cctNumber   : AData[n] := cell^.NumberValue;
              cctDateTime : AData[n] := cell^.DateTimeValue
            end;
            inc(n);
            if n = Length(AData) then SetLength(AData, length(AData) + BLOCKSIZE);
          end;
        end
    else
    if (arg.ResultType in [rtInteger, rtFloat, rtDateTime, rtCell, rtBoolean]) then
    begin
      AData[n] := ArgToFloat(arg);
      inc(n);
      if n = Length(AData) then SetLength(AData, Length(AData) + BLOCKSIZE);
    end;
  end;
  SetLength(AData, n);
end;


{------------------------------------------------------------------------------}
{   Conversion of simple data types to ExpressionResults                          }
{------------------------------------------------------------------------------}

function BooleanResult(AValue: Boolean): TsExpressionResult;
begin
  Result.ResultType := rtBoolean;
  Result.ResBoolean := AValue;
end;

function CellResult(AValue: String): TsExpressionResult;
begin
  Result.ResultType := rtCell;
  ParseCellString(AValue, Result.ResRow, Result.ResCol);
end;

function CellResult(ACellRow, ACellCol: Cardinal): TsExpressionResult;
begin
  Result.ResultType := rtCell;
  Result.ResRow := ACellRow;
  Result.ResCol := ACellCol;
end;

function DateTimeResult(AValue: TDateTime): TsExpressionResult;
begin
  Result.ResultType := rtDateTime;
  Result.ResDateTime := AValue;
end;

function EmptyResult: TsExpressionResult;
begin
  Result.ResultType := rtEmpty;
end;

function ErrorResult(const AValue: TsErrorValue): TsExpressionResult;
begin
  Result.ResultType := rtError;
  Result.ResError := AValue;
end;

function FloatResult(const AValue: TsExprFloat): TsExpressionResult;
begin
  Result.ResultType := rtFloat;
  Result.ResFloat := AValue;
end;

function IntegerResult(const AValue: Integer): TsExpressionResult;
begin
  Result.ResultType := rtInteger;
  Result.ResInteger := AValue;
end;

function IsInteger(const AValue: TsExpressionResult): Boolean;
var
  i: Int64;
  cell: PCell;
begin
  Result := false;
  case AValue.ResultType of
    rtString : Result := TryStrToInt64(AValue.ResString, i);
    rtInteger: Result := true;
    rtFloat  : Result := (frac(AValue.ResFloat) = 0);
    rtEmpty  : Result := true;
    rtCell   : begin
                 cell := AValue.Worksheet.FindCell(AValue.ResRow, AValue.ResCol);
                 if Assigned(cell) then
                   case cell^.ContentType of
                     cctNumber:
                       Result := frac(cell^.NumberValue) = 0.0;
                     cctDateTime:
                       Result := frac(cell^.DateTimeValue) = 0.0;
                     cctUTF8String:
                       Result := TryStrToInt64(cell^.UTF8StringValue, i);
                   end;
               end;
  end;
end;

function IsString(const AValue: TsExpressionResult): Boolean;
var
  cell: PCell;
begin
  Result := false;
  case AValue.ResultType of
    rtString: Result := true;
    rtCell  : begin
                cell := AValue.Worksheet.FindCell(AValue.ResRow, AValue.ResCol);
                Result := (cell <> nil) and (cell^.ContentType = cctUTF8String);
              end;
  end;
end;

function StringResult(const AValue: string): TsExpressionResult;
begin
  Result.ResultType := rtString;
  Result.ResString := AValue;
end;

{@@ ---------------------------------------------------------------------------
  Registers a non-built-in function:

  @param AName       Name of the function as used for calling it in the spreadsheet
  @param AResultType A character classifying the data type of the function result:
                       'I' integer
                       'F' floating point number
                       'D' date/time value
                       'S' string
                       'B' boolean value (TRUE/FALSE)
                       'R' cell range, can also be used for functions requiring
                           a cell "reference", like "CELL(..)"
  @param AParamTypes A string with result type symbols for each parameter of the
                     function. Symbols as used for "ResultType" with these
                     additions:
                       - Use a lower-case character if a parameter is optional.
                         (must be at the end of the string)
                       - Add "+" if the last parameter type is valid for a variable
                         parameter count (Excel does pose a limit of 30, though).
                       - Use "?" if the data type should not be checked.
  @param AExcelCode  ID of the function needed in the xls biff file. Please see
                     the "OpenOffice Documentation of Microsoft Excel File Format"
                     section 3.11.
  @param ACallBack   Address of the procedure called when the formula is
                     calculated.
-------------------------------------------------------------------------------}
procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallback: TsExprFunctionCallBack);
begin
  with BuiltinIdentifiers do
    AddFunction(bcUser, AName, AResultType, AParamTypes, AExcelCode, ACallBack);
end;

procedure RegisterFunction(const AName: ShortString; const AResultType: Char;
  const AParamTypes: String; const AExcelCode: Integer; ACallback: TsExprFunctionEvent);
begin
  with BuiltinIdentifiers do
    AddFunction(bcUser, AName, AResultType, AParamTypes, AExcelCode, ACallBack);
end;


{ TsBuiltInExprIdentifierDef }

procedure TsBuiltInExprIdentifierDef.Assign(Source: TPersistent);
begin
  inherited Assign(Source);
  if Source is TsBuiltInExprIdentifierDef then
    FCategory := (Source as TsBuiltInExprIdentifierDef).Category;
end;

initialization
  ExprFormatSettings := DefaultFormatSettings;
  ExprFormatSettings.DecimalSeparator := '.';
  ExprFormatSettings.ListSeparator := ',';

  RegisterStdBuiltins(BuiltinIdentifiers);

finalization
  FreeBuiltins;

end.
