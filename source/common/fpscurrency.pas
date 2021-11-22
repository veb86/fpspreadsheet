{@@ ----------------------------------------------------------------------------
  Unit **fpsCurrency** provides a variety of **utility functions** related to
  currency values.
  
  Currency symbols are the key to detection of currency values. 
  In order to reckognize strings as currency symbols they have to be registered 
  in the internal CurrencyList.

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}

unit fpsCurrency;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

procedure RegisterCurrency(ACurrencySymbol: String);
procedure RegisterCurrencies(AList: TStrings; AReplace: Boolean);
procedure UnregisterCurrency(ACurrencySymbol: String);
function  CurrencyRegistered(ACurrencySymbol: String): Boolean;
procedure GetRegisteredCurrencies(AList: TStrings);

function IsNegative(var AText: String): Boolean;
function RemoveCurrencySymbol(ACurrencySymbol: String;
  var AText: String): Boolean;
function TryStrToCurrency(AText: String; out ANumber: Double;
  out ACurrencySymbol:String; const AFormatSettings: TFormatSettings): boolean;


implementation

var
  CurrencyList: TStrings = nil;

{@@ ----------------------------------------------------------------------------
  Registers a currency symbol UTF8 string for usage by fpspreadsheet

  Currency symbols are the key for detection of currency values. In order to
  reckognize strings as currency symbols they have to be registered in the
  internal CurrencyList.

  Registration occurs automatically for USD, "$", the currencystring defined
  in the DefaultFormatSettings and for the currency symbols used explicitly
  when calling WriteCurrency or WriteNumerFormat.
  
  @param    ACurrencySymbol   A string symbol which identifies the currency, e.g. '$'                           
-------------------------------------------------------------------------------}
procedure RegisterCurrency(ACurrencySymbol: String);
begin
  if not CurrencyRegistered(ACurrencySymbol) and (ACurrencySymbol <> '') then
    CurrencyList.Add(ACurrencySymbol);
end;

{@@ ----------------------------------------------------------------------------
  RegisterCurrencies() registers the currency strings contained in the string list
  If AReplace is true, the list replaces the currently registered list.
  
  @param   AList     A list containing the strings to be registered as currency symbols.
  @param   AReplace  When this is @true the currently registered symbols are deleted, otherwise the strings are registered as additional currency symbols.
-------------------------------------------------------------------------------}
procedure RegisterCurrencies(AList: TStrings; AReplace: Boolean);
var
  i: Integer;
begin
  if AList = nil then
    exit;

  if AReplace then CurrencyList.Clear;
  for i:=0 to AList.Count-1 do
    RegisterCurrency(AList[i]);
end;

{@@ ----------------------------------------------------------------------------
  Removes registration of a currency symbol string for usage by fpspreadsheet
  
  @param    ACurrencySymbol   The currency symbol which is to be un-registered.
-------------------------------------------------------------------------------}
procedure UnregisterCurrency(ACurrencySymbol: String);
var
  i: Integer;
begin
  i := CurrencyList.IndexOf(ACurrencySymbol);
  if i <> -1 then CurrencyList.Delete(i);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a string is registered as valid currency symbol string

  @param    ACurrencySymbol   The currency symbol to be checked.
  @returns  @true of the currency symbol is registered, @false otherwise.
-------------------------------------------------------------------------------}
function CurrencyRegistered(ACurrencySymbol: String): Boolean;
begin
  Result := CurrencyList.IndexOf(ACurrencySymbol) <> -1;
end;

{@@ ----------------------------------------------------------------------------
  Writes all registered currency symbols to a string list

  @param    AList   List to which all the registered currency symbols are copied.
-------------------------------------------------------------------------------}
procedure GetRegisteredCurrencies(AList: TStrings);
begin
  AList.Clear;
  AList.Assign(CurrencyList);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given number string is a negative value. In case of
  currency values, this can be indicated by brackets, or a minus sign at string
  start or end.
  
  @param    AText  String to be analyzed. Brackets or minus sign will be removed if the string is found to correspond to a negative value.
  @returns  @true when the input string was detected to be a negative number, @false otherwise.
-------------------------------------------------------------------------------}
function IsNegative(var AText: String): Boolean;
begin
  Result := false;
  if AText = '' then
    exit;
  if (AText[1] = '(') and (AText[Length(AText)] = ')') then
  begin
    Result := true;
    Delete(AText, 1, 1);
    Delete(AText, Length(AText), 1);
    AText := Trim(AText);
  end else
  if (AText[1] = '-') then
  begin
    Result := true;
    Delete(AText, 1, 1);
    AText := Trim(AText);
  end else
  if (AText[Length(AText)] = '-') then
  begin
    Result := true;
    Delete(AText, Length(AText), 1);
    AText := Trim(AText);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks wheter a specified currency symbol is contained in a string, removes
  the currency symbol and returns the remaining string.
  
  @param    ACurrencySymbol   The currency symbol to be detected.
  @param    AText             String to be checked. On output, the currency symbol will be removed from the string if found.
  @returns  @true if the input string did contain a currency symbol, @false otherwise.
-------------------------------------------------------------------------------}
function RemoveCurrencySymbol(ACurrencySymbol: String; var AText: String): Boolean;
var
  p: Integer;
begin
  p := pos(ACurrencySymbol, AText);
  if p > 0 then
  begin
    Delete(AText, p, Length(ACurrencySymbol));
    AText := Trim(AText);
    Result := true;
  end else
    Result := false;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a string is a number with attached currency symbol. Looks also
  for negative values in brackets.
  
  @param  AText    String which is to be converted to a currency value.
  @param  ANumber  Returns the numeric value of the string
  @param  ACurrencySymbol  Returns the currency symbol found.
  @param  AFormatSettings  Contains the decimal separator used in the input string.
  @returns @true if the string could be converted to a number successfully, @false otherwise.
-------------------------------------------------------------------------------}
function TryStrToCurrency(AText: String; out ANumber: Double;
  out ACurrencySymbol:String; const AFormatSettings: TFormatSettings): boolean;
var
  i: Integer;
  s: String;
  isNeg: Boolean;
begin
  Result := false;
  ANumber := 0.0;
  ACurrencySymbol := '';

  // Check the text for the presence of each known curreny symbol
  for i:= 0 to CurrencyList.Count-1 do
  begin
    // Store string in temporary variable since it will be modified
    s := AText;
    // Check for this currency sign being contained in the string, remove it if found.
    if RemoveCurrencySymbol(CurrencyList[i], s) then
    begin
      // Check for negative signs and remove them, but keep this information
      isNeg := IsNegative(s);
      // Try to convert remaining string to number
      if TryStrToFloat(s, ANumber, AFormatSettings) then begin
        // if successful: take care of negative values
        if isNeg then ANumber := -ANumber;
        ACurrencySymbol := CurrencyList[i];
        Result := true;
        exit;
      end;
    end;
  end;
end;

initialization
  // Known currency symbols
  CurrencyList := TStringList.Create;
  with TStringList(CurrencyList) do
  begin
    CaseSensitive := false;
    Duplicates := dupIgnore;
  end;
  RegisterCurrency('USD');
  RegisterCurrency('$');
  RegisterCurrency(AnsiToUTF8(DefaultFormatSettings.CurrencyString));

finalization
  FreeAndNil(CurrencyList);

end.

