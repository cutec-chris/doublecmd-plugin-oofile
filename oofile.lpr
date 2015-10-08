library oofile;

{$mode objfpc}{$H+}
{$include calling.inc}

uses
  Classes, sysutils, WLXPlugin, fprichdocument, oodocument, laz_fpspreadsheet,
  fpolestorage, uminiconvencoding, general_nogui, Utils;

procedure ListGetDetectString(DetectString:pchar;maxlen:integer); dcpcall;
begin
  StrCopy(DetectString, 'EXT="ODG"|EXT="ODT"|EXT="ODS"|EXT="DOC"|EXT="TXT"|EXT="ME"|EXT="LST"|EXT="FCSTD"');
end;

function StripUnwantedChar(Text: string):string;
var
  Allowed: Set of Char;
  i, LeftOvers: Integer;
begin
  Allowed := [' ', '0'..'9', 'a'..'z', 'A'..'Z', '~'..')', '-', '.', '\', ':', '`', '/', '<', ',', '>', ';', '{', '}',#13,#9];

  SetLength(Result, Length(Text));
  LeftOvers := 1;
  for i := 1 to Length(Text) do begin
    if Text[i] in Allowed then begin
      Result[LeftOvers]:= Text[i];
      Inc(LeftOvers);
    end
  end;
  SetLength(Result, LeftOvers-1);
end;
function GetWordText(aFileName: string; var aText: string
  ): Boolean;
var
  MemStream: TMemoryStream;
  OLEStorage: TOLEStorage;
  OLEDocument : TOLEDocument;
  aStringStream: TStringStream;
  aContent : string;
  aContent2: String;
begin
  MemStream := TMemoryStream.Create;
  OLEStorage := TOLEStorage.Create;
  try
    // Only one stream is necessary for any number of worksheets
    OLEDocument.Stream := MemStream;
    OLEStorage.ReadOLEFile(aFileName, OLEDocument,'WordDocument');
    if MemStream.Seek($800,soFromBeginning) = $800 then
      begin
        Setlength(aContent,MemStream.Size-$800);
        MemStream.Read(aContent[1],MemStream.Size-$800);
        aContent2 := ConvertEncoding(aContent,EncodingUCS2LE,EncodingUTF8);
        aText:=StripUnwantedChar(aContent2);
      end;
  finally
    OLEStorage.Free;
  end;
end;


function ListGetText(FileToLoad:pchar;contentbuf:pchar;contentbuflen:integer):pchar; dcpcall;
var
  aDoc: TODFDocument;
  aFile: string;
  aText: TStringList;
  bText: string;
  i: Integer;
begin
  aFile := FileToLoad;
  if FileExists(aFile) then
    begin
      aText := TStringList.Create;
      case lowercase(ExtractFileExt(aFile)) of
      '.doc':
          begin
            GetWordText(aFile,bText);
            aText.Text:=bText;
          end;
       '.odt','.ods':
           begin
             aDoc := TODFDocument.Create;
             aDoc.FileName:=FileToLoad;
             aDoc.Open;
             aText.Text := aDoc.AsString;
             aDoc.Free;
           end
      else
        begin
          aText.LoadFromFile(aFile);
          for i := 0 to 1500 do
            if (length(copy(aText.Text,i,1))>0) and (ord(copy(aText.Text,i,1)[1]) > 127) then
              begin
                break;
              end;
           aText.Text := copy(aText.Text,0,1500);
        end;
      end;
      if aText.Count>0 then
        begin
          Result := PChar(aText.Text);
        end;
    end;
  aText.Free;
end;

function ListGetPreviewBitmapFile(FileToLoad:pchar;OutputPath:pchar;width,height:integer;
    contentbuf:pchar;contentbuflen:integer):pchar; dcpcall;
var
  aDoc: TODFDocument;
begin
  Result := '';
  case lowercase(ExtractFileExt(FileToLoad)) of
   '.odt','.ods','.odg','.fcstd':
     begin
       aDoc := TODFDocument.Create;
       try
         aDoc.FileName:=FileToLoad;
         aDoc.Open;
         if aDoc.ExtractFile('Thumbnails/thumbnail.png',OutputPath) then
           Result := PChar(OutputPath+'Thumbnails'+DirectorySeparator+'thumbnail.png');
         aDoc.Free;
       except
       end;
     end;
  end;
end;

exports
  ListGetDetectString,
  ListGetText,
  ListGetPreviewBitmapFile;

begin
end.

