unit FRCUnit.Settings;

interface

uses
  System.Classes, System.JSON, System.SysUtils;

type
  TFRCSettings = class(TPersistent)
  strict private
    FReadOnlyDirectories: TStringList;
    FReadWriteDirectories: TStringList;
    function DirectoryListToDelimitedString(const AList: TStringList): string;
    function SaveToJSONString: string;
    procedure LoadFromJSONString(const AJSON: string);
    procedure ParseDirectoriesIntoList(const ASemicolonSeparatedDirectories: string; const AList: TStringList;
      const AClearList: Boolean);
    procedure SetReadOnlyDirectories(const AValue: TStringList);
    procedure SetReadWriteDirectories(const AValue: TStringList);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Assign(ASource: TPersistent); override;

    function ReadOnlyDirectoriesAsString: string;
    function ReadWriteDirectoriesAsString: string;
    procedure Clear;
    procedure LoadFromFile(const AFileName: string);
    procedure ParseReadOnlyDirectoriesFromString(const ASemicolonSeparatedDirectories: string; const AClearList: Boolean = True);
    procedure ParseReadWriteDirectoriesFromString(const ASemicolonSeparatedDirectories: string; const AClearList: Boolean = True);
    procedure SaveToFile(const AFileName: string);
  published
    property ReadOnlyDirectories: TStringList read FReadOnlyDirectories write SetReadOnlyDirectories;
    property ReadWriteDirectories: TStringList read FReadWriteDirectories write SetReadWriteDirectories;
  end;

implementation

const
  SETTING_NAME_READONLY_DIRECTORIES = 'ReadOnlyDirectories';
  SETTING_NAME_READWRITE_DIRECTORIES = 'ReadWriteDirectories';

function StringListToJSONArray(const AList: TStringList): TJSONArray;
begin
  Result := TJSONArray.Create;

  for var LStringValue in AList do
    Result.Add(LStringValue);
end;

// Fills a TStringList from a TJSONArray
procedure JSONArrayToStringList(const AArray: TJSONArray; const AList: TStringList);
begin
  AList.Clear;

  if not Assigned(AArray) then
    Exit;

  for var LItem in AArray do
    AList.Add(LItem.Value);
end;

{ TFRCSettings }

procedure TFRCSettings.Clear;
begin
  FReadOnlyDirectories.Clear;
  FReadWriteDirectories.Clear;
end;

constructor TFRCSettings.Create;
begin
  inherited Create;

  FReadOnlyDirectories  := TStringList.Create;
  FReadOnlyDirectories.Delimiter := ';';
  FReadOnlyDirectories.StrictDelimiter := True;

  FReadWriteDirectories := TStringList.Create;
  FReadWriteDirectories.Delimiter := ';';
  FReadWriteDirectories.StrictDelimiter := True;
end;

destructor TFRCSettings.Destroy;
begin
  FReadOnlyDirectories.Free;
  FReadWriteDirectories.Free;

  inherited Destroy;
end;

procedure TFRCSettings.ParseReadOnlyDirectoriesFromString(const ASemicolonSeparatedDirectories: string; const AClearList: Boolean = True);
begin
  ParseDirectoriesIntoList(ASemicolonSeparatedDirectories, FReadOnlyDirectories, AClearList);
end;

procedure TFRCSettings.ParseReadWriteDirectoriesFromString(const ASemicolonSeparatedDirectories: string; const AClearList: Boolean = True);
begin
  ParseDirectoriesIntoList(ASemicolonSeparatedDirectories, FReadWriteDirectories, AClearList);
end;

procedure TFRCSettings.ParseDirectoriesIntoList(const ASemicolonSeparatedDirectories: string; const AList: TStringList;
  const AClearList: Boolean);
begin
  if AClearList then
    AList.Clear;

  var LDirectories := ASemicolonSeparatedDirectories.Split([';']);

  for var LDirectory in LDirectories do
    AList.Add(IncludeTrailingPathDelimiter(LDirectory));
end;

procedure TFRCSettings.Assign(ASource: TPersistent);
begin
  if ASource is TFRCSettings then
  begin
    var LSrc := TFRCSettings(ASource);

    FReadOnlyDirectories.Assign(LSrc.FReadOnlyDirectories);
    FReadWriteDirectories.Assign(LSrc.FReadWriteDirectories);
  end
  else
    inherited Assign(ASource);
end;

procedure TFRCSettings.SetReadOnlyDirectories(const AValue: TStringList);
begin
  FReadOnlyDirectories.Assign(AValue);
end;

procedure TFRCSettings.SetReadWriteDirectories(const AValue: TStringList);
begin
  FReadWriteDirectories.Assign(AValue);
end;

function TFRCSettings.SaveToJSONString: string;
begin
  var LRoot := TJSONObject.Create;
  try
    LRoot.AddPair(SETTING_NAME_READONLY_DIRECTORIES,  StringListToJSONArray(FReadOnlyDirectories));
    LRoot.AddPair(SETTING_NAME_READWRITE_DIRECTORIES, StringListToJSONArray(FReadWriteDirectories));

    Result := LRoot.Format; // pretty-printed. Use LRoot.ToJSON for compact format
  finally
    LRoot.Free;
  end;
end;

procedure TFRCSettings.LoadFromJSONString(const AJSON: string);
const
  EXCEPTION_MESSAGE = 'Invalid JSON for TFRCSettings';
begin
  var LRoot := TJSONObject.ParseJSONValue(AJSON) as TJSONObject;

  if not Assigned(LRoot) then
    raise EArgumentException.Create(EXCEPTION_MESSAGE);

  try
    JSONArrayToStringList(LRoot.GetValue<TJSONArray>(SETTING_NAME_READONLY_DIRECTORIES), FReadOnlyDirectories);
    JSONArrayToStringList(LRoot.GetValue<TJSONArray>(SETTING_NAME_READWRITE_DIRECTORIES), FReadWriteDirectories);
  finally
    LRoot.Free;
  end;
end;

function TFRCSettings.ReadOnlyDirectoriesAsString: string;
begin
  Result := DirectoryListToDelimitedString(FReadOnlyDirectories);
end;

procedure TFRCSettings.SaveToFile(const AFileName: string);
begin
  var LSL := TStringList.Create;
  try
    LSL.Text := SaveToJSONString;
    LSL.SaveToFile(AFileName, TEncoding.UTF8);
  finally
    LSL.Free;
  end;
end;

function TFRCSettings.DirectoryListToDelimitedString(const AList: TStringList): string;
begin
  Result := '';

  if AList.Count = 0 then
    Exit;

  for var LIndex := 0 to AList.Count - 2 do
    Result := Result + AList[LIndex] + AList.Delimiter;

  Result := Result + AList[AList.Count - 1];
end;

procedure TFRCSettings.LoadFromFile(const AFileName: string);
begin
  Clear;

  if not FileExists(AFileName) then
    Exit;

  var LSL := TStringList.Create;
  try
    LSL.LoadFromFile(AFileName, TEncoding.UTF8);
    LoadFromJSONString(LSL.Text);
  finally
    LSL.Free;
  end;
end;

function TFRCSettings.ReadWriteDirectoriesAsString: string;
begin
  Result := DirectoryListToDelimitedString(FReadWriteDirectories);
end;

end.
