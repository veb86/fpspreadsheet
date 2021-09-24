program GuiTestProject;

{$mode objfpc}{$H+}

uses
  Interfaces, Forms, GuiTestRunner, 
  ReadFieldsTestUnit, SortTestUnit, SearchTestUnit, FilterTestUnit, PostTestUnit,
  EmptyColumnsTestUnit, CopyFromDatasetUnit;

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TGuiTestRunner, TestRunner);
  Application.Run;
end.

