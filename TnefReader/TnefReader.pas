unit TnefReader;

interface

uses
  Classes, SysUtils, TnefAttributeConsts;

type
  TTnefReader = class
  private
    FInputStream: TStream;
    FAttributeValueStart: Int64;
    FChecksum: SmallInt;
    FAttributeLength: Integer;
    FAttributeTag: TTnefAttributeTag;
    FAttachmentKey: SmallInt;
    FAttributeLevel: TTnefAttributeLevel;
    FTnefVersion: Integer;
    FOemCodePage: Integer;

    function GetRawReadPosition: Int64;
    function GetAttributeType: TTnefAttributeType;
    function ReadBuffer(ALength: Integer): TBytes;
    procedure UpdateCheckSum(const ABuffer: TBytes);
    function ReadAttributeLevel: TTnefAttributeLevel;
    procedure ReadAttributeValue;
    procedure ReadHeader;
    procedure ValidateNextAttributePos;
  public
    constructor Create(AInputStream: TStream);

    function GetEncoding: string;
    function NextAttribute: Boolean;
    procedure ValidateChecksum;

    function ReadByte: Byte;
    function ReadInt16: SmallInt;
    function ReadInt32: Integer;
    function ReadBytes(ALength: Integer): TBytes; overload;
    procedure ReadBytes(ADestination: TStream; ALength: Integer); overload;
    function ReadString: string;

    property AttachmentKey: SmallInt read FAttachmentKey;
    property AttributeLevel: TTnefAttributeLevel read FAttributeLevel;
    property AttributeTag: TTnefAttributeTag read FAttributeTag;
    property AttributeLength: Integer read FAttributeLength;
    property TnefVersion: Integer read FTnefVersion;
    property OemCodePage: Integer read FOemCodePage;
    property RawReadPosition: Int64 read GetRawReadPosition;
    property AttributeType: TTnefAttributeType read GetAttributeType;
  end;

implementation

uses
  TnefAttribute, clTranslator;

const
  TnefSignature = $223e9f78;

{ TTnefReader }

function TTnefReader.ReadBytes(ALength: Integer): TBytes;
begin
  Result := ReadBuffer(ALength);
  UpdateCheckSum(Result);
end;

constructor TTnefReader.Create(AInputStream: TStream);
begin
  inherited Create();

  FInputStream := AInputStream;

  FAttributeLevel := alMessage;
  FAttributeTag := agNull;
  FAttributeLength := 0;
  FAttributeValueStart := 0;

  FTnefVersion := 0;
  FOemCodePage := 0;

  FChecksum := 0;
  FAttachmentKey := 0;

  ReadHeader();
end;

function TTnefReader.GetAttributeType: TTnefAttributeType;
begin
  Result := GetTnefAttributeType(TnefAttributeTags[AttributeTag] and $F0000);
end;

type
  TclCharSetCodePage = record
    CodePage: Integer;
    Name: string[20];
  end;

const
  CharSetCodePages: array [0..34] of TclCharSetCodePage = (
    (CodePage: 1250; Name: 'windows-1250'),
    (CodePage: 1251; Name: 'windows-1251'),
    (CodePage: 1252; Name: 'windows-1252'),
    (CodePage: 1253; Name: 'windows-1253'),
    (CodePage: 1254; Name: 'windows-1254'),
    (CodePage: 1255; Name: 'windows-1255'),
    (CodePage: 1256; Name: 'windows-1256'),
    (CodePage: 1257; Name: 'windows-1257'),
    (CodePage: 1258; Name: 'windows-1258'),
    (CodePage: 28591; Name: 'iso-8859-1'),
    (CodePage: 28592; Name: 'iso-8859-2'),
    (CodePage: 28593; Name: 'iso-8859-3'),
    (CodePage: 28594; Name: 'iso-8859-4'),
    (CodePage: 28595; Name: 'iso-8859-5'),
    (CodePage: 28596; Name: 'iso-8859-6'),
    (CodePage: 28597; Name: 'iso-8859-7'),
    (CodePage: 28598; Name: 'iso-8859-8'),
    (CodePage: 28599; Name: 'iso-8859-9'),
    (CodePage: 28603; Name: 'iso-8859-13'),
    (CodePage: 28605; Name: 'iso-8859-15'),
    (CodePage: 866; Name: 'ibm866'),
    (CodePage: 866; Name: 'cp866'),
    (CodePage: 1200; Name: 'utf-16'),
    (CodePage: 12000; Name: 'utf-32'),
    (CodePage: 65000; Name: 'utf-7'),
    (CodePage: 65001; Name: 'utf-8'),
    (CodePage: 20127; Name: 'us-ascii'),
    (CodePage: 28591; Name: 'Latin1'),
    (CodePage: 10007; Name: 'x-mac-cyrillic'),
    (CodePage: 21866; Name: 'koi8-u'),
    (CodePage: 20866; Name: 'koi8-r'),
    (CodePage: 932; Name: 'shift-jis'),
    (CodePage: 932; Name: 'shift_jis'),
    (CodePage: 50220; Name: 'iso-2022-jp'),
    (CodePage: 50220; Name: 'csISO2022JP')
  );

function TTnefReader.GetEncoding: string;
var
  i: Integer;
begin
  for i := Low(CharSetCodePages) to High(CharSetCodePages) do
  begin
    if (CharSetCodePages[i].CodePage = OemCodePage) then
    begin
      Result := CharSetCodePages[i].Name;
      Exit;
    end;
  end;
  Result := '';
end;

function TTnefReader.GetRawReadPosition: Int64;
begin
  Result := FInputStream.Position;
end;

function TTnefReader.NextAttribute: Boolean;
begin
  ValidateNextAttributePos();

  if (RawReadPosition >= FInputStream.Size) then
  begin
    Result := False;
    Exit;
  end;

  FAttributeLevel := ReadAttributeLevel();

  FAttributeTag := GetTnefAttributeTag(ReadInt32());

  FAttributeLength := ReadInt32();
  FAttributeValueStart := RawReadPosition;

  FChecksum := 0;

  ReadAttributeValue();

  Result := True;
end;

function TTnefReader.ReadAttributeLevel: TTnefAttributeLevel;
begin
  Result := GetTnefAttributeLevel(ReadByte());
end;

procedure TTnefReader.ReadAttributeValue;
var
  versionAttribute: TTnefVersionAttribute;
  codePageAttribute: TTnefOemCodePageAttribute;
begin
  if (AttributeLevel <> alMessage) then Exit;

  case (AttributeTag) of
    agTnefVersion:
    begin
      versionAttribute := TTnefVersionAttribute.Create(Self);
      try
        versionAttribute.Load();
        FTnefVersion := versionAttribute.TnefVersion;
      finally
        versionAttribute.Free();
      end;
    end;
    agOemCodepage:
    begin
      codePageAttribute := TTnefOemCodePageAttribute.Create(Self);
      try
        codePageAttribute.Load();
        FOemCodePage := codePageAttribute.OemCodePage;
      finally
        codePageAttribute.Free();
      end;
    end;
  end;
end;

function TTnefReader.ReadBuffer(ALength: Integer): TBytes;
var
  readLength: Integer;
begin
  SetLength(Result, ALength);
  readLength := FInputStream.Read(Result[0], ALength);

  if (readLength <> ALength) then
  begin
    raise Exception.Create('Invalid Stream');
  end;
end;

function TTnefReader.ReadByte: Byte;
var
  buffer: TBytes;
begin
  buffer := ReadBuffer(1);
  UpdateCheckSum(buffer);
  Result := buffer[0];
end;

procedure TTnefReader.ReadBytes(ADestination: TStream; ALength: Integer);
const
  MaxBufSize = 4096;
var
  bufSize, n: Integer;
  buf: TBytes;
begin
  bufSize := ALength;
  if (bufSize > MaxBufSize) then
  begin
    bufSize := MaxBufSize;
  end;

  while (ALength <> 0) do
  begin
    n := bufSize;
    if (n > ALength) then
    begin
      n := ALength;
    end;

    buf := ReadBytes(n);
    ADestination.Write(buf[0], n);
    Dec(ALength, n);
  end;
end;

procedure TTnefReader.ReadHeader;
var
  signature: Integer;
begin
  signature := ReadInt32();

  if (signature <> TnefSignature) then
  begin
    raise Exception.Create('Invalid TNEF format');
  end;

  FAttachmentKey := ReadInt16();
end;

function TTnefReader.ReadInt16: SmallInt;
var
  index: Integer;
  buffer: TBytes;
begin
  index := 0;
  buffer := ReadBuffer(2);
  UpdateCheckSum(buffer);

  Result := SmallInt(buffer[index] or (buffer[index + 1] shl 8));
end;

function TTnefReader.ReadInt32: Integer;
var
  index: Integer;
  buffer: TBytes;
begin
  index := 0;
  buffer := ReadBuffer(4);
  UpdateCheckSum(buffer);

  Result := buffer[index] or (buffer[index + 1] shl 8) or (buffer[index + 2] shl 16) or (buffer[index + 3] shl 24);
end;

function TTnefReader.ReadString: string;
var
  bytes: TBytes;
  len: Integer;
begin
  bytes := ReadBytes(AttributeLength);

  len := Length(bytes);
  if (len > 0) and (bytes[len - 1] = 0) then
  begin
    Dec(len);
  end;

  Result := TclTranslator.GetString(bytes, 0, len, GetEncoding());
end;

procedure TTnefReader.UpdateCheckSum(const ABuffer: TBytes);
var
  i: Integer;
begin
  for i := 0 to Length(ABuffer) - 1 do
  begin
    FChecksum := SmallInt((FChecksum + ABuffer[i]) and $FFFF);
  end;
end;

procedure TTnefReader.ValidateChecksum;
var
  nextAttributePos: Int64;
  real, etalon: SmallInt;
begin
  nextAttributePos := FAttributeValueStart + AttributeLength;
  if (nextAttributePos = RawReadPosition) then
  begin
    real := FChecksum;
    etalon := ReadInt16();

    if (real <> etalon) then
    begin
      raise Exception.Create('Invalid checksum');
    end;
  end;
end;

procedure TTnefReader.ValidateNextAttributePos;
var
  nextAttributePos: Int64;
begin
  if (FAttributeValueStart = 0) then Exit;

  nextAttributePos := FAttributeValueStart + AttributeLength + 2;
  if (nextAttributePos > RawReadPosition) then
  begin
    FInputStream.Seek(nextAttributePos, soBeginning);
    if (RawReadPosition > FInputStream.Size) then
    begin
      raise Exception.Create('Invalid stream');
    end;
  end else
  if (nextAttributePos < RawReadPosition) then
  begin
    raise Exception.Create('The attribute was read incorrectly');
  end;
end;

end.
