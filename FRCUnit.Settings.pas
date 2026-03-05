unit FRCUnit.Settings;

interface

uses
  System.Classes, System.JSON, System.SysUtils, REST.Json;

type
  TFRCSettings = class(TPersistent)
  strict private
    FReadOnlyDirectories: TStringList;
    FReadWriteDirectories: TStringList;
    function ListToString(const AList: TStringList): string;
    procedure AddSemicolonSepatratedDirectories(const ASemicolonSepaaratedDirectories: string; const AList: TStringList);
    procedure SetReadOnlyDirectories(const AValue: TStringList);
    procedure SetReadWriteDirectories(const AValue: TStringList);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Assign(ASource: TPersistent); override;

    procedure Clear;

    procedure AddReadOnlyDirectories(const ASemicolonSepaaratedDirectories: string);
    procedure AddReadWriteDirectories(const ASemicolonSepaaratedDirectories: string);
    function ReadOnlyDirectoriesStr: string;
    function ReadWriteDirectoriesStr: string;
    // Serialization
    function ToJSON: string;
    procedure SaveToFile(const AFileName: string);
    procedure LoadFromFile(const AFileName: string);
    procedure FromJSON(const AJSON: string);
  published
    property ReadOnlyDirectories: TStringList read FReadOnlyDirectories write SetReadOnlyDirectories;
    property ReadWriteDirectories: TStringList read FReadWriteDirectories write SetReadWriteDirectories;
  end;

implementation


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

  FReadWriteDirectories := TStringList.Create;
  FReadWriteDirectories.Delimiter := ';';
end;

destructor TFRCSettings.Destroy;
begin
  FReadOnlyDirectories.Free;
  FReadWriteDirectories.Free;

  inherited Destroy;
end;

procedure TFRCSettings.AddReadOnlyDirectories(const ASemicolonSepaaratedDirectories: string);
begin
  AddSemicolonSepatratedDirectories(ASemicolonSepaaratedDirectories, FReadOnlyDirectories);
end;

procedure TFRCSettings.AddReadWriteDirectories(const ASemicolonSepaaratedDirectories: string);
begin
  AddSemicolonSepatratedDirectories(ASemicolonSepaaratedDirectories, FReadWriteDirectories);
end;

procedure TFRCSettings.AddSemicolonSepatratedDirectories(const ASemicolonSepaaratedDirectories: string; const AList: TStringList);
begin
  var LDirectories := ASemicolonSepaaratedDirectories.Split([';']);

  for var LDirectory in LDirectories do
  begin
    if DirectoryExists(LDirectory) then
      AList.Add(IncludeTrailingPathDelimiter(LDirectory));
  end;
end;

procedure TFRCSettings.Assign(ASource: TPersistent);
var
  LSrc: TFRCSettings;
begin
  if ASource is TFRCSettings then
  begin
    LSrc := TFRCSettings(ASource);

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

function StringListToJSONArray(const AList: TStringList): TJSONArray;
var
  LStringValue: string;
begin
  Result := TJSONArray.Create;

  for LStringValue in AList do
    Result.Add(LStringValue);
end;

// Fills a TStringList from a TJSONArray
procedure JSONArrayToStringList(const AArray: TJSONArray; const AList: TStringList);
var
  LItem: TJSONValue;
begin
  AList.Clear;

  if not Assigned(AArray) then
    Exit;

  for LItem in AArray do
    AList.Add(LItem.Value);
end;

function TFRCSettings.ToJSON: string;
var
  LRoot: TJSONObject;
begin
  LRoot := TJSONObject.Create;
  try
    LRoot.AddPair('ReadOnlyDirectories',  StringListToJSONArray(FReadOnlyDirectories));
    LRoot.AddPair('ReadWriteDirectories', StringListToJSONArray(FReadWriteDirectories));

    Result := LRoot.Format; // pretty-printed; use .ToJSON for compact
  finally
    LRoot.Free;
  end;
end;

procedure TFRCSettings.FromJSON(const AJSON: string);
var
  LRoot: TJSONObject;
begin
  LRoot := TJSONObject.ParseJSONValue(AJSON) as TJSONObject;

  if not Assigned(LRoot) then
    raise EArgumentException.Create('Invalid JSON for TFRCSettings');

  try
    JSONArrayToStringList(LRoot.GetValue<TJSONArray>('ReadOnlyDirectories'), FReadOnlyDirectories);
    JSONArrayToStringList(LRoot.GetValue<TJSONArray>('ReadWriteDirectories'), FReadWriteDirectories);
  finally
    LRoot.Free;
  end;
end;

function TFRCSettings.ReadOnlyDirectoriesStr: string;
begin
  Result := ListToString(FReadOnlyDirectories);
end;

procedure TFRCSettings.SaveToFile(const AFileName: string);
var
  LSL: TStringList;
begin
  LSL := TStringList.Create;
  try
//     if FileExists(AFileName) then
//       if not DeleteFile(AFileName) then
//         raise EFilerError.Create('File could not be deleted: ' + AFileName.QuotedString('"'));

    LSL.Text := ToJSON;
    LSL.SaveToFile(AFileName, TEncoding.UTF8);
  finally
    LSL.Free;
  end;
end;

function TFRCSettings.ListToString(const AList: TStringList): string;
begin
  Result := '';

  if AList.Count = 0 then
    Exit;

  for var LIndex := 0 to AList.Count - 2 do
    Result := Result + AList[LIndex] + AList.Delimiter;

  Result := Result + AList[AList.Count - 1];
end;

procedure TFRCSettings.LoadFromFile(const AFileName: string);
var
  LSL: TStringList;
begin
  Clear;

  if not FileExists(AFileName) then
    Exit;

  LSL := TStringList.Create;
  try
    LSL.LoadFromFile(AFileName, TEncoding.UTF8);
    FromJSON(LSL.Text);
  finally
    LSL.Free;
  end;
end;

function TFRCSettings.ReadWriteDirectoriesStr: string;
begin
  Result := ListToString(FReadWriteDirectories);
end;

end.
