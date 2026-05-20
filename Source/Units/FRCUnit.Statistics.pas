unit FRCUnit.Statistics;

// Result types produced by TFileRightsChecker — the error taxonomy, the per-error
// record, the collection that holds them, and the simple run-level counters.
// Kept separate from the checker so consumers (forms, reports, tests) can depend
// only on these data types without pulling the full WinAPI surface of the checker.

interface

uses
  System.Generics.Collections, System.SysUtils;

type
  TFileRightErrorType = (
    frcNone,
    frcMissingPrivilege,
    frcFileNotReadable,
    frcFileNotWritable,
    frcUserHasNoExecuteRightsForFile,
    frcIsReparsePoint,
    frcUACVirtualization,
    frcDirectoryNotReadable,
    frcDirectoryNotWritable,
    frcReadOnlyAttribute,
    frcOwnershipMismatch,
    frcEffectiveRightsMissing,
    frcEFSEncrypted,
    frcNetworkShare,
    frcShareModeConflict,
    frcEmptyDACL,
    frcExplicitDenyACE);

  TErrorItem = class(TObject)
  strict private
    FFileSystemItem: string;
    FErrorType: TFileRightErrorType;
    FErrorDescription: string;
    function GetErrorTypeStr: string;
  public
    constructor Create(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string);

    property FileSystemItem: string read FFileSystemItem;
    property ErrorType: TFileRightErrorType read FErrorType;
    property ErrorTypeStr: string read GetErrorTypeStr;
    property ErrorDescription: string read FErrorDescription;
  end;

  TErrorItemCollection = class(TObject)
  strict private
    FErrorItems: TObjectList<TErrorItem>;
    function GetItem(const AIndex: Integer): TErrorItem;
  public
    constructor Create;
    destructor Destroy; override;

    function Count: Integer;
    procedure Add(const AItem: TErrorItem);
    property Items[const AIndex: Integer]: TErrorItem read GetItem; default;
  end;

  TStatistics = class(TObject)
  strict private
    FFilesChecked: Integer;
    FDirectoriesChecked: Integer;
  public
    procedure AddCheckedFile;
    procedure AddCheckedDirectory;
    property FilesChecked: Integer read FFilesChecked;
    property DirectoriesChecked: Integer read FDirectoriesChecked;
  end;

implementation

{ TErrorItem }

constructor TErrorItem.Create(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string);
begin
  inherited Create;

  FFileSystemItem := AFileSystemItem;
  FErrorType := AErrorType;
  FErrorDescription := AErrorDescription;
end;

function TErrorItem.GetErrorTypeStr: string;
begin
  case FErrorType of
    frcNone: Result := '';
    frcMissingPrivilege: Result := 'Process is missing required privilege';
    frcFileNotReadable: Result := 'File not readable';
    frcFileNotWritable: Result := 'File not writable';
    frcUserHasNoExecuteRightsForFile: Result := 'User has no execute rights';
    frcIsReparsePoint: Result := 'Path is a reparse point or symlink';
    frcUACVirtualization: Result := 'Path is under UAC virtualization';
    frcDirectoryNotReadable: Result := 'Directory not readable';
    frcDirectoryNotWritable: Result := 'Directory not writable';
    frcReadOnlyAttribute: Result := 'Read-only attribute set';
    frcOwnershipMismatch: Result := 'Current user is not the owner';
    frcEffectiveRightsMissing: Result := 'Effective rights missing for current user';
    frcEFSEncrypted: Result := 'File or directory is EFS-encrypted';
    frcNetworkShare: Result := 'Path is on a network share (SMB / mapped drive)';
    frcShareModeConflict: Result := 'File open blocked by sharing mode';
    frcEmptyDACL: Result := 'Empty DACL — all access denied';
    frcExplicitDenyACE: Result := 'Explicit DENY ACE applies to current user';
    else
      Result := Format('Unknown error type (%d)', [Ord(FErrorType)]);
  end;
end;

{ TErrorItemCollection }

procedure TErrorItemCollection.Add(const AItem: TErrorItem);
begin
  FErrorItems.Add(AItem);
end;

function TErrorItemCollection.Count: Integer;
begin
  Result := FErrorItems.Count;
end;

constructor TErrorItemCollection.Create;
begin
  inherited Create;

  FErrorItems := TObjectList<TErrorItem>.Create(True);
end;

destructor TErrorItemCollection.Destroy;
begin
  FErrorItems.Free;

  inherited Destroy;
end;

function TErrorItemCollection.GetItem(const AIndex: Integer): TErrorItem;
begin
  Result := FErrorItems[AIndex];
end;

{ TStatistics }

procedure TStatistics.AddCheckedDirectory;
begin
  Inc(FDirectoriesChecked);
end;

procedure TStatistics.AddCheckedFile;
begin
  Inc(FFilesChecked);
end;

end.
