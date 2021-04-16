unit fpsvisualreg;

{$MODE objfpc}{$H+}
{$DEFINE REGISTER_ALL_FILE_FORMATS}

{$R ../../resource/fpsvisualreg.res}

interface

uses
  Classes, SysUtils;

procedure Register;
  
implementation

uses
  LResources, ActnList, PropEdits,
  {$IFDEF REGISTER_ALL_FILE_FORMATS}
  fpsallformats,
  {$ENDIF}
  fpspreadsheetctrls, fpspreadsheetgrid, fpspreadsheetchart, fpsactions;
  
{@@ ----------------------------------------------------------------------------
  Registers the visual spreadsheet components in the Lazarus component palette,
  page "FPSpreadsheet".
-------------------------------------------------------------------------------}
procedure Register;
begin
  RegisterComponents('FPSpreadsheet', [
    TsWorkbookSource,
    TsWorkbookTabControl, TsWorksheetIndicator,
    TsWorksheetGrid,
    TsCellEdit, TsCellIndicator, TsCellCombobox,
    TsSpreadsheetInspector
  ]);

  RegisterComponents('Chart', [
    TsWorksheetChartSource,
    TsWorkbookChartSource
  ]);

  RegisterActions('FPSpreadsheet', [
    // Worksheet-releated actions
    TsWorksheetAddAction, TsWorksheetDeleteAction, TsWorksheetRenameAction,
    TsWorksheetZoomAction,
    // Cell or cell range formatting actions
    TsCopyAction,
    TsClearFormatAction,
    TsFontStyleAction, TsFontDialogAction, TsBackgroundColorDialogAction,
    TsHorAlignmentAction, TsVertAlignmentAction,
    TsTextRotationAction, TsWordWrapAction,
    TsNumberFormatAction, TsDecimalsAction,
    TsCellProtectionAction,
    TsCellBorderAction, TsNoCellBordersAction,
    TsCellCommentAction, TsCellHyperlinkAction,
    TsMergeAction
  ], nil);

  RegisterPropertyEditor(TypeInfo(TFileName),
    TsWorkbookSource, 'FileName', TFileNamePropertyEditor
  );

end;

initialization
  RegisterPropertyToSkip(TsSpreadsheetInspector, 'RowHeights',
    'For compatibility with older Laz versions.', '');

  RegisterPropertyToSkip(TsSpreadsheetInspector, 'ColWidths',
    'For compatibility with older Laz versions.', '');

end.

