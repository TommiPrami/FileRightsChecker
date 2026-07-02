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
    frcExplicitDenyACE,
    frcCloudPlaceholder,
    frcPathTooLong,
    frcReadOnlyVolume,
    frcNonNTFSFileSystem,
    frcLowDiskSpace,
    frcControlledFolderAccess,
    frcExclusiveOpenFailed,
    frcFileReservedByProcess,
    frcFileNoDeleteRights,
    frcFileNoWriteDACRights);

  // Findings are not all equal: an in-use file or a network-share note is context,
  // not a failure. Errors are things that actually block the customer application.
  TErrorSeverity = (esInfo, esWarning, esError);

  TErrorItem = class(TObject)
  strict private
    FFileSystemItem: string;
    FErrorType: TFileRightErrorType;
    FErrorDescription: string;
    FSeverity: TErrorSeverity;
    function GetErrorTypeStr: string;
    function GetSeverityStr: string;
  public
    constructor Create(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string;
      const ASeverity: TErrorSeverity = esError);

    property FileSystemItem: string read FFileSystemItem;
    property ErrorType: TFileRightErrorType read FErrorType;
    property ErrorTypeStr: string read GetErrorTypeStr;
    property ErrorDescription: string read FErrorDescription;
    property Severity: TErrorSeverity read FSeverity;
    property SeverityStr: string read GetSeverityStr;
  end;

  TErrorItemCollection = class(TObject)
  strict private
    FErrorItems: TObjectList<TErrorItem>;
    function GetItem(const AIndex: Integer): TErrorItem;
  public
    constructor Create;
    destructor Destroy; override;

    function Count: Integer;
    function CountBySeverity(const ASeverity: TErrorSeverity): Integer;
    function CountByErrorType(const AErrorType: TFileRightErrorType): Integer;
    // Count of esError items only — what the progress display calls "Errors".
    function ErrorCount: Integer;
    procedure Add(const AItem: TErrorItem);
    property Items[const AIndex: Integer]: TErrorItem read GetItem; default;
  end;

  TStatistics = class(TObject)
  strict private
    FFilesChecked: Integer;
    FDirectoriesChecked: Integer;
    FFilesOpenedExclusively: Integer;
  public
    procedure AddCheckedFile;
    procedure AddCheckedDirectory;
    procedure AddCheckedExclusiveFile;
    property FilesChecked: Integer read FFilesChecked;
    property DirectoriesChecked: Integer read FDirectoriesChecked;
    property FilesOpenedExclusively: Integer read FFilesOpenedExclusively;
  end;

implementation

{ TErrorItem }

constructor TErrorItem.Create(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string;
  const ASeverity: TErrorSeverity = esError);
begin
  inherited Create;

  FFileSystemItem := AFileSystemItem;
  FErrorType := AErrorType;
  FErrorDescription := AErrorDescription;
  FSeverity := ASeverity;
end;

function TErrorItem.GetSeverityStr: string;
begin
  case FSeverity of
    esInfo: Result := 'INFO';
    esWarning: Result := 'WARNING';
    esError: Result := 'ERROR';
    else
      Result := Format('Unknown severity (%d)', [Ord(FSeverity)]);
  end;
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
    frcCloudPlaceholder: Result := 'File is offline or a cloud placeholder';
    frcPathTooLong: Result := 'Path exceeds MAX_PATH';
    frcReadOnlyVolume: Result := 'Volume is read-only';
    frcNonNTFSFileSystem: Result := 'File system has no NTFS ACLs';
    frcLowDiskSpace: Result := 'Low free disk space';
    frcControlledFolderAccess: Result := 'Defender Controlled Folder Access is active';
    frcExclusiveOpenFailed: Result := 'File cannot be opened in exclusive mode';
    frcFileReservedByProcess: Result := 'File is held open by another process';
    frcFileNoDeleteRights: Result := 'User has no DELETE right on file';
    frcFileNoWriteDACRights: Result := 'User has no WRITE_DAC right on file';
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

function TErrorItemCollection.CountBySeverity(const ASeverity: TErrorSeverity): Integer;
begin
  Result := 0;

  for var LItem in FErrorItems do
    if LItem.Severity = ASeverity then
      Inc(Result);
end;

function TErrorItemCollection.CountByErrorType(const AErrorType: TFileRightErrorType): Integer;
begin
  Result := 0;

  for var LItem in FErrorItems do
    if LItem.ErrorType = AErrorType then
      Inc(Result);
end;

function TErrorItemCollection.ErrorCount: Integer;
begin
  Result := CountBySeverity(esError);
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

procedure TStatistics.AddCheckedExclusiveFile;
begin
  Inc(FFilesOpenedExclusively);
end;

end.
