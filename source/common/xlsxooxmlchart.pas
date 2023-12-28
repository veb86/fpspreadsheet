unit xlsxooxmlChart;

{$mode objfpc}{$H+}
{$include ..\fps.inc}

interface

{$ifdef FPS_CHARTS}

uses
  Classes, SysUtils, StrUtils, Contnrs, FPImage,
  {$ifdef FPS_PATCHED_ZIPPER}fpszipper,{$else}zipper,{$endif}
  laz2_xmlread, laz2_DOM,
  fpsTypes, fpSpreadsheet, fpsChart, fpsUtils, fpsNumFormat,
  fpsReaderWriter, fpsXMLCommon;

type

  { TsSpreadOOXMLChartReader }

  TsSpreadOOXMLChartReader = class(TsBasicSpreadChartReader)
  private
    FPointSeparatorSettings: TFormatSettings;

  protected

  public
    constructor Create(AReader: TsBasicSpreadReader); override;
    destructor Destroy; override;

  end;

  TsSpreadOOXMLChartWriter = class(TsBasicSpreadChartWriter)
  private
    FPointSeparatorSettings: TFormatSettings;

  protected

  public
    constructor Create(AWriter: TsBasicSpreadWriter); override;
    destructor Destroy; override;

  end;

{$ENDIF}

implementation

{$IFDEF FPS_CHARTS}

uses
  xlsxooxml;

{ TsSpreadOOXMLChartReader }

constructor TsSpreadOOXMLChartReader.Create(AReader: TsBasicSpreadReader);
begin
  inherited Create(AReader);

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';
end;

destructor TsSpreadOOXMLChartReader.Destroy;
begin
  inherited;
end;


{ TsSpreadOOXMLChartWriter }

constructor TsSpreadOOXMLChartWriter.Create(AWriter: TsBasicSpreadWriter);
begin
  inherited Create(AWriter);

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';
end;

destructor TsSpreadOOXMLChartWriter.Destroy;
begin
  inherited;
end;

{$ENDIF}

end.

