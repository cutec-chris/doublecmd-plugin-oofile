library oofile;

{$mode objfpc}{$H+}
{$include calling.inc}

uses
  Classes,
  sysutils,
  WLXPlugin, fprichdocument,oodocument;

procedure ListGetDetectString(DetectString:pchar;maxlen:integer); dcpcall;
begin
  StrCopy(DetectString, 'EXT="ODT"');
end;

function ListGetText(FileToLoad:pchar;contentbuf:pchar;contentbuflen:integer):pchar; dcpcall;
begin
end;

function ListGetPreviewBitmapFile(FileToLoad:pchar;OutputPath:pchar;width,height:integer;
    contentbuf:pchar;contentbuflen:integer):pchar; dcpcall;
var
  aDoc: TODFDocument;
begin
  aDoc := TODFDocument.Create;
  aDoc.FileName:=FileToLoad;
  aDoc.Open;

  aDoc.Free;
end;

exports
  ListGetDetectString,
  ListGetText,
  ListGetPreviewBitmapFile;

begin
end.

