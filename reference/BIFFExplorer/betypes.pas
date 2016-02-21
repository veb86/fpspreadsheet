unit beTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  BIFFNODE_TXO_CONTINUE1 = 1;
  BIFFNODE_TXO_CONTINUE2 = 2;
  BIFFNODE_SST_CONTINUE  = 3;

type
  { Virtual tree node data }
  TBiffNodeData = record //class
    Offset: Integer;
    RecordID: Integer;
    RecordName: String;
    RecordDescription: String;
    Index: Integer;
    Tag: Integer;
  end;
  PBiffNodeData = ^TBiffNodeData;

implementation

end.

