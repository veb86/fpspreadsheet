program spreadtestgui;

{$mode objfpc}{$H+}

{.$DEFINE HEAPTRC}    // Instead of using -gh activate this to write the heap trace to file

uses
 {$IFDEF HEAPTRC}
  SysUtils,
 {$ENDIF}
  Interfaces, Forms, GuiTestRunner, testsutility, datetests, stringtests,
  numberstests, manualtests, internaltests, mathtests, fileformattests,
  formattests, colortests, fonttests, optiontests, conditionalformattests,
  numformatparsertests,
  formulatests, rpnFormulaUnit, singleformulatests, calcformulatests,
  exceltests, emptycelltests, errortests, virtualmodetests, colrowtests,
  ssttests, celltypetests, sortingtests, copytests, movetests, enumeratortests,
  commenttests, hyperlinktests, pagelayouttests, protectiontests,
  definednames_tests;

begin
 {$IFDEF HEAPTRC}
  // Assuming your build mode sets -dDEBUG in Project Options/Other when defining -gh
  // This avoids interference when running a production/default build without -gh

  if FileExists('heap.trc') then
    DeleteFile('heap.trc');
  SetHeapTraceOutput('heap.trc');
 {$ENDIF HEAPTRC}

  Application.Initialize;
  Application.CreateForm(TGuiTestRunner, TestRunner);
  Application.Run;
end.

