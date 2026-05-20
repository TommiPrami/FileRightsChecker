unit FRCUnit.FileRightsChecker;

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Generics.Collections, FRCUnit.WinAPI;

type
  TFileRightErrorType = (frcNone, frcMissingPrivilege, frcFileNotReadable, frcFileNotWritable, frcUserHasNoExecuteRightsForFile,
    frcIsReparsePoint, frcUACVirtualization, frcDirectoryNotReadable, frcDirectoryNotWritable,
    frcReadOnlyAttribute, frcOwnershipMismatch, frcEffectiveRightsMissing, frcEFSEncrypted, frcNetworkShare,
    frcShareModeConflict, frcEmptyDACL, frcExplicitDenyACE);

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

  TFileRightsChecker = class(TObject)
  strict private
    FErrors: TErrorItemCollection;
    FReadOnlyStatistics: TStatistics;
    FReadWriteStatistics: TStatistics;
    FOpenFilesLongFileAndPathNameSupport: Boolean;
    FCheckProcessBackupPrivileges: Boolean;
    FRunDirectoryGetEffectiveRightsShortfallTests: Boolean;
    FRunFileGetEffectiveRightsShortfallTests: Boolean;
    FRunCurrentUserIsOwnerTests: Boolean;
    function GetTempFileName(const ADirectory: string): string;
    function HasExplicitDenyACE(const ADirectory: string; var AErrorDescription: string): Boolean;
    function TestDirectoryReadRights(const ADirectory: string; var AErrorDescription: string): Boolean;
    function TestDirectoryWriteRights(const ADirectory: string; var AErrorDescription: string): Boolean;
    function TestExecuteRights(const AFileName: string; var AErrorDescription: string): Boolean;
    function TestOpenFileRights(const AFileName: string; const AInReadWriteMode: Boolean; var AErrorDescription: string): Boolean;
    function IsRunningElevated(var AErrorDescription: string): Boolean;
    function ToLongPath(const ADirectory: string): string;
    function HasEmptyDACL(const ADirectory: string; var AErrorDescription: string): Boolean;
    function IsDirectoryUnderUACVirtualization(const ADirectory: string; var AErrorDescription: string): Boolean;
    function HasPrivilege(const APrivilegeName: string; var AErrorDescription: string): Boolean;
    function IsReparsePoint(const ADirectory: string; var AErrorDescription: string): Boolean;
    function GetFileIntegrityLevel(const ADirectory: string; var AErrorDescription: string): string;
    function IsDirectoryEmpty(const ADirectory: string): Boolean;
    function DiagnoseFileShareModes(const AFileName: string; const AInReadWriteMode: Boolean; var AErrorDescription: string): Boolean;
    function HasReadOnlyAttribute(const APath: string; var AErrorDescription: string): Boolean;
    function IsEFSEncrypted(const APath: string; var AErrorDescription: string): Boolean;
    function IsOnNetworkShare(const APath: string; var AErrorDescription: string): Boolean;
    function CurrentUserIsOwner(const APath: string; var AErrorDescription: string): Boolean;
    function GetEffectiveRightsShortfall(const APath: string; const ACheckWriteRights: Boolean; var AErrorDescription: string): Boolean;
    procedure CheckToAddMoreInfoForCreateFileFailure(const AErrorCode: DWORD; var AErrorDescription: string);
    procedure GetExceptionErrorDescription(const AErrorMethod, AFileSystemItem: string; const AException: Exception; var AErrorDescription: string);
    procedure GetFilesAndDirs(const ADirectory: string; const AFiles, ADirectories: TStringList;  var AErrorDescription: string;
      const AClearLists: Boolean = True);
    procedure LogError(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string);
    procedure CheckProcessBackupPrivileges(const ADirectory: string);
    procedure InitializeDirectoriesAndFiles(const ADirectory: string; const AFiles, ASubDirectories: TStringList; var AErrorDescription: string);
    procedure DoDirectoryChecks(const ADirectories: TStringList; const ACheckWriteRights: Boolean);
    procedure DoFileChecks(const AFiles: TStringList; const ACheckWriteRights: Boolean);
  public
    constructor Create(const AOpenFilesLongFileAndPathNameSupport: Boolean = True; const ACheckProcessBackupPrivileges: Boolean = False;
      const ARunDirectoryGetEffectiveRightsShortfallTests: Boolean = False; const ARunFileGetEffectiveRightsShortfallTests: Boolean = False;
      const ARunCurrentUserIsOwnerTests: Boolean = False);
    destructor Destroy; override;

    procedure Execute(const ADirectory: string; const ACheckWriteRights: Boolean);
    property Errors: TErrorItemCollection read FErrors;
    property ReadWriteStatistics: TStatistics read FReadWriteStatistics;
    property ReadOnlyStatistics: TStatistics read FReadOnlyStatistics;
    property RunDirectoryGetEffectiveRightsShortfallTests: Boolean read FRunDirectoryGetEffectiveRightsShortfallTests write FRunDirectoryGetEffectiveRightsShortfallTests;
    property RunFileGetEffectiveRightsShortfallTests: Boolean read FRunFileGetEffectiveRightsShortfallTests write FRunFileGetEffectiveRightsShortfallTests;
    property RunCurrentUserIsOwnerTests: Boolean read FRunCurrentUserIsOwnerTests write FRunCurrentUserIsOwnerTests;
  end;

implementation

uses
  System.Math, System.StrUtils;

function IsProcess64Bit: Boolean;
begin
  Result := SizeOf(Pointer) = 8;
end;

// Returns True only if the SID denotes the current effective user or one of their groups.
// Failing the API call counts as "unknown" -> False, so we don't fire on unrelated SIDs.
function CurrentTokenIsMemberOf(const ASid: PSID): Boolean;
var
  LIsMember: BOOL;
begin
  Result := False;

  if ASid = nil then
    Exit;

  if not IsValidSid(ASid) then
    Exit;

  LIsMember := False;

  // Passing 0 for the token tells CheckTokenMembership to use the impersonation
  // token of the calling thread, or the process token if none — i.e. "us".
  if CheckTokenMembership(0, ASid, LIsMember) then
    Result := LIsMember;
end;

{ TFileRightsChecker }

procedure TFileRightsChecker.InitializeDirectoriesAndFiles(const ADirectory: string; const AFiles, ASubDirectories: TStringList;
  var AErrorDescription: string);
begin
  AErrorDescription := '';

  var LDirectories := ADirectory.Split([';']);

  for var LDirectory in LDirectories do
  begin
    var LTrimmed := LDirectory.Trim;

    if LTrimmed.IsEmpty then
      Continue;

    if not DirectoryExists(LTrimmed) then
    begin
      LogError(LTrimmed, frcDirectoryNotReadable, 'Top-level directory does not exist or is not accessible');
      Continue;
    end;

    var LPerDirError: string := '';
    GetFilesAndDirs(LTrimmed, AFiles, ASubDirectories, LPerDirError, False);

    if not LPerDirError.IsEmpty then
      LogError(LTrimmed, frcDirectoryNotReadable, LPerDirError);

    ASubDirectories.Insert(0, LTrimmed);

    if FCheckProcessBackupPrivileges then
      CheckProcessBackupPrivileges(LTrimmed);
  end;
end;

function TFileRightsChecker.IsDirectoryEmpty(const ADirectory: string): Boolean;
var
  LSearchRec: TSearchRec;
begin
  Result := True;

  if FindFirst(IncludeTrailingPathDelimiter(ADirectory) + '*', faAnyFile, LSearchRec) = 0 then
  try
    repeat
      if (LSearchRec.Name <> '.') and (LSearchRec.Name <> '..') then
      begin
        Result := False;
        Break;
      end;
    until FindNext(LSearchRec) <> 0;
  finally
    FindClose(LSearchRec);
  end;
end;

function TFileRightsChecker.HasEmptyDACL(const ADirectory: string; var AErrorDescription: string): Boolean;
var
  LBytesNeeded: DWORD;
  LSecDesc: PSECURITY_DESCRIPTOR;
  LDACL: PACL;
  LDACLPresent: BOOL;
  LDefaulted: BOOL;
begin
  Result := False;
  AErrorDescription := '';
  LBytesNeeded := 0;

  try
    // First call to get required buffer size
    GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION, nil, 0, LBytesNeeded);

    if LBytesNeeded = 0 then
    begin
      AErrorDescription := Format('GetFileSecurity failed [%d]: %s',
        [GetLastError, SysErrorMessage(GetLastError)]);
      Exit;
    end;

    LSecDesc := AllocMem(NativeInt(LBytesNeeded));
    try
      if not GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION,
         LSecDesc, LBytesNeeded, LBytesNeeded) then
      begin
        AErrorDescription := Format('GetFileSecurity failed [%d]: %s', [GetLastError, SysErrorMessage(GetLastError)]);

        Exit;
      end;

      if not GetSecurityDescriptorDacl(LSecDesc, LDACLPresent, LDACL, LDefaulted) then
      begin
        AErrorDescription := Format('GetSecurityDescriptorDacl failed [%d]: %s', [GetLastError, SysErrorMessage(GetLastError)]);

        Exit;
      end;

      // LDACLPresent=True but LDACL=nil means empty DACL = deny everyone
      if LDACLPresent and (LDACL = nil) then
      begin
        AErrorDescription := 'Empty DACL detected — all access denied to everyone including Administrators';

        Result := True;
      end;
    finally
      FreeMem(LSecDesc);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('HasEmptyDACL', ADirectory, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.IsRunningElevated(var AErrorDescription: string): Boolean;
var
  LTokenHandle: THandle;
  LElevation: TOKEN_ELEVATION;
  LReturnLength: DWORD;
begin
  Result := False;

  try
    if OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, LTokenHandle) then
    try
      if GetTokenInformation(LTokenHandle, TokenElevation, @LElevation, SizeOf(LElevation), LReturnLength) then
        Result := LElevation.TokenIsElevated <> 0;
    finally
      CloseHandle(LTokenHandle);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('IsRunningElevated', '', E, AErrorDescription);
  end;
end;

function TFileRightsChecker.HasExplicitDenyACE(const ADirectory: string; var AErrorDescription: string): Boolean;
var
  LBytesNeeded: DWORD;
  LSecDesc: PSECURITY_DESCRIPTOR;
  LDACL: PACL;
  LDACLPresent: BOOL;
  LDefaulted: BOOL;
  LAceIndex: DWORD;
  LAceHeader: PACE_HEADER;
  LAccessDeniedAce: PACCESS_DENIED_ACE;
  LAclSizeInfo: TACL_SIZE_INFORMATION;
begin
  Result := False;
  AErrorDescription := '';
  LBytesNeeded := 0;

  try
    // First call to get required buffer size
    GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION, nil, 0, LBytesNeeded);

    if LBytesNeeded = 0 then
    begin
      AErrorDescription := Format('GetFileSecurity failed [%d]: %s', [GetLastError, SysErrorMessage(GetLastError)]);

      Exit;
    end;

    LSecDesc := AllocMem(NativeInt(LBytesNeeded));
    try
      // Second call to get the actual security descriptor
      if not GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION, LSecDesc, LBytesNeeded, LBytesNeeded) then
      begin
        AErrorDescription := Format('GetFileSecurity failed [%d]: %s', [GetLastError, SysErrorMessage(GetLastError)]);

        Exit;
      end;

      if not GetSecurityDescriptorDacl(LSecDesc, LDACLPresent, LDACL, LDefaulted) then
      begin
        AErrorDescription := Format('GetSecurityDescriptorDacl failed [%d]: %s', [GetLastError, SysErrorMessage(GetLastError)]);

        Exit;
      end;

      // No DACL present means full access to everyone — not a deny situation
      if not LDACLPresent or (LDACL = nil) then
        Exit;

      // Get number of ACEs in the DACL
      if not GetAclInformation(LDACL^, @LAclSizeInfo, SizeOf(LAclSizeInfo), TAclInformationClass(AclSizeInformation)) then
      begin
        AErrorDescription := Format('GetAclInformation failed [%d]: %s', [GetLastError, SysErrorMessage(GetLastError)]);

        Exit;
      end;

      for LAceIndex := 0 to LAclSizeInfo.AceCount - 1 do
      begin
        if not GetAce(LDACL^, LAceIndex, Pointer(LAceHeader)) then
          Continue;

        if LAceHeader^.AceType <> ACCESS_DENIED_ACE_TYPE then
          Continue;

        // INHERIT_ONLY ACEs are templates for child objects and do not apply to
        // the object whose ACL we're inspecting.
        if (LAceHeader^.AceFlags and INHERIT_ONLY_ACE_FLAG) <> 0 then
          Continue;

        LAccessDeniedAce := PACCESS_DENIED_ACE(LAceHeader);

        // Only flag if the DENY actually applies to the current process token —
        // otherwise we'd report deny ACEs for unrelated SIDs (Guests, SYSTEM, etc.)
        // that don't affect us.
        if not CurrentTokenIsMemberOf(@LAccessDeniedAce^.SidStart) then
          Continue;

        // Get the SID name for reporting who is denied
        var LSIDName: array[0..255] of Char;
        var LDomainName: array[0..255] of Char;
        var LSIDNameLen: DWORD := Length(LSIDName);
        var LDomainNameLen: DWORD := Length(LDomainName);
        var LSIDNameUse: SID_NAME_USE;

        if LookupAccountSid(nil, @LAccessDeniedAce^.SidStart, LSIDName, LSIDNameLen, LDomainName, LDomainNameLen, LSIDNameUse) then
          AErrorDescription := AErrorDescription + Format('DENY ACE found for: %s\%s  ', [LDomainName, LSIDName])
        else
          AErrorDescription := AErrorDescription + 'DENY ACE found for unknown SID  ';

        Result := True;
      end;

    finally
      FreeMem(LSecDesc);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('HasExplicitDenyACE', ADirectory, E, AErrorDescription);
  end;
end;

procedure TFileRightsChecker.GetFilesAndDirs(const ADirectory: string; const AFiles, ADirectories: TStringList;
  var AErrorDescription: string; const AClearLists: Boolean = True);
begin
  try
    if AClearLists then
    begin
      AFiles.Clear;
      ADirectories.Clear;
    end;

    var LPath := IncludeTrailingPathDelimiter(ADirectory);

    var LSearchRec: TSearchRec;
    if FindFirst(LPath + '*', faAnyFile, LSearchRec) = 0 then
    try
      repeat
        if (LSearchRec.Name = '.') or (LSearchRec.Name = '..') then
          Continue;

        if (LSearchRec.Attr and faDirectory) <> 0 then
        begin
          ADirectories.Add(LPath + LSearchRec.Name);

          // Don't follow reparse points (junctions, symlinks) — they can loop or
          // expand into unrelated trees. The caller still sees the directory in
          // the list so file-level checks can flag it as a reparse point.
{$WARN SYMBOL_PLATFORM OFF}
          if (LSearchRec.Attr and faSymLink) = 0 then
            GetFilesAndDirs(LPath + LSearchRec.Name, AFiles, ADirectories, AErrorDescription, False);  // recurse
{$WARN SYMBOL_PLATFORM ON}
        end
        else
          AFiles.Add(LPath + LSearchRec.Name);

      until FindNext(LSearchRec) <> 0;
    finally
      FindClose(LSearchRec);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('GetFilesAndDirs', ADirectory, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.GetTempFileName(const ADirectory: string): string;

  function CleanGUID(const AGUIDStr: string): string;
  begin
    Result := AGUIDStr;

    Result := StringReplace(Result, '{', '', [rfReplaceAll]);
    Result := StringReplace(Result, '}', '', [rfReplaceAll]);
  end;

var
  LGUID: TGUID;
begin
  CreateGUID(LGUID);

  Result := IncludeTrailingPathDelimiter(ADirectory) + 'FileRightsChecker_' + CleanGUID(GUIDToString(LGUID)) + '.tmp';
end;

procedure TFileRightsChecker.LogError(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string);
begin
  var LErrorItem := TErrorItem.Create(AFileSystemItem, AErrorType, AErrorDescription);

  Errors.Add(LErrorItem);
end;

function TFileRightsChecker.TestDirectoryWriteRights(const ADirectory: string; var AErrorDescription: string): Boolean;
var
  LTempFile: string;
  LFileHandle: THandle;
  LData: AnsiString;
  LBytesWritten: DWORD;
  LErrorCode: DWORD;
begin
  Result := False;
  AErrorDescription := '';
  LTempFile := GetTempFileName(ADirectory);

  try
    // --- Create / Open ---
    LFileHandle := CreateFile(PChar(ToLongPath(LTempFile)), GENERIC_WRITE, 0, nil, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0);

    if LFileHandle = INVALID_HANDLE_VALUE then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('CreateFile failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

      CheckToAddMoreInfoForCreateFileFailure(LErrorCode, AErrorDescription);

      Exit;
    end;

    try
      // --- Write ---
      LData := 'FileRightsChecker test data';

      if not WriteFile(LFileHandle, LData[1], DWORD(Length(LData)), LBytesWritten, nil) or (LBytesWritten <> DWORD(Length(LData))) then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('WriteFile failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

        Exit;
      end;
    finally
      CloseHandle(LFileHandle);
    end;

    // --- Delete ---
    if not DeleteFile(PChar(ToLongPath(LTempFile))) then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('DeleteFile failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);
      Exit;
    end;

    Result := True;
  except
    on E: Exception do
      GetExceptionErrorDescription('TestDirectoryWriteRights', ADirectory, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.TestDirectoryReadRights(const ADirectory: string; var AErrorDescription: string): Boolean;
var
  LDirHandle: THandle;
  LErrorCode: DWORD;
begin
  Result := False;
  AErrorDescription := '';

  try
    LDirHandle := CreateFile(PChar(ToLongPath(ADirectory)), GENERIC_READ, FILE_SHARE_READ or FILE_SHARE_WRITE, nil, OPEN_EXISTING,
      FILE_FLAG_BACKUP_SEMANTICS { required to open a directory handle },  0);

    if LDirHandle = INVALID_HANDLE_VALUE then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('Directory read rights check failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

      Exit;
    end;

    CloseHandle(LDirHandle);
    Result := True;
  except
    on E: Exception do
      GetExceptionErrorDescription('TestDirectoryReadRights', ADirectory, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.TestExecuteRights(const AFileName: string; var AErrorDescription: string): Boolean;
var
  LFileHandle: THandle;
  LErrorCode: DWORD;
begin
  Result := False;
  AErrorDescription := '';

  try
    if not FileExists(AFileName) then
    begin
      AErrorDescription := Format('File not found: %s', [AFileName]);
      Exit;
    end;

    LFileHandle := CreateFile(PChar(ToLongPath(AFileName)), GENERIC_EXECUTE, FILE_SHARE_READ, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
    try
      if LFileHandle = INVALID_HANDLE_VALUE then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('Execute rights check failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

        Exit;
      end;

      Result := True;
    finally
      if LFileHandle <> INVALID_HANDLE_VALUE then
        CloseHandle(LFileHandle);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('TestExecuteRights', AFileName, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.TestOpenFileRights(const AFileName: string; const AInReadWriteMode: Boolean;
  var AErrorDescription: string): Boolean;
var
  LFileHandle: THandle;
  LAccessMode: DWORD;
  LErrorCode: DWORD;
begin
  Result := False;
  AErrorDescription := '';

  try
    if AInReadWriteMode then
      LAccessMode := GENERIC_READ or GENERIC_WRITE
    else
      LAccessMode := GENERIC_READ;

    // Permissive share mode — this is a passive probe, we shouldn't fail just because
    // another process has the file open for writing or pending delete.
    LFileHandle := CreateFile(PChar(ToLongPath(AFileName)), LAccessMode,
      FILE_SHARE_READ or FILE_SHARE_WRITE or FILE_SHARE_DELETE,
      nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

    if LFileHandle = INVALID_HANDLE_VALUE then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('OpenFile failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

      Exit;
    end;

    CloseHandle(LFileHandle);

    Result := True;
  except
    on E: Exception do
      GetExceptionErrorDescription('TestOpenFileRights', AFileName, E, AErrorDescription);
  end;
end;

// Probes the file with several share modes separately. The most-permissive mode
// is the same one TestOpenFileRights uses, so success/failure at that level is
// already reflected there. The point of this function is to surface info from
// the RESTRICTIVE modes: if a more restrictive mode fails (sharing violation),
// another process has the file open and the customer's application may also see
// intermittent open failures even with sufficient ACL rights.
//
// Returns True if any restrictive mode reported a non-permission failure that
// looks like a sharing conflict.
function TFileRightsChecker.DiagnoseFileShareModes(const AFileName: string; const AInReadWriteMode: Boolean;
  var AErrorDescription: string): Boolean;

  function ProbeShareMode(const AShareMode: DWORD; const ALabel: string; const AAccessMode: DWORD;
    var ADetails: string): Boolean;
  var
    LHandle: THandle;
    LErr: DWORD;
  begin
    LHandle := CreateFile(PChar(ToLongPath(AFileName)), AAccessMode, AShareMode, nil, OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL, 0);

    if LHandle = INVALID_HANDLE_VALUE then
    begin
      LErr := GetLastError;
      ADetails := ADetails + Format(' | %s: failed [%d] %s', [ALabel, LErr, SysErrorMessage(LErr)]);
      // ERROR_SHARING_VIOLATION (32) is the smoking gun for share-mode conflicts.
      Result := LErr = ERROR_SHARING_VIOLATION;
    end
    else
    begin
      CloseHandle(LHandle);
      ADetails := ADetails + Format(' | %s: ok', [ALabel]);
      Result := False;
    end;
  end;

var
  LAccessMode: DWORD;
  LDetails: string;
  LSharingViolation: Boolean;
begin
  Result := False;
  AErrorDescription := '';
  LDetails := '';
  LSharingViolation := False;

  if AInReadWriteMode then
    LAccessMode := GENERIC_READ or GENERIC_WRITE
  else
    LAccessMode := GENERIC_READ;

  try
    // Each share value documents what OTHER openers we'll tolerate while we have the handle.
    LSharingViolation := ProbeShareMode(0, 'exclusive (no share)', LAccessMode, LDetails) or LSharingViolation;
    LSharingViolation := ProbeShareMode(FILE_SHARE_READ, 'share-read', LAccessMode, LDetails) or LSharingViolation;
    LSharingViolation := ProbeShareMode(FILE_SHARE_READ or FILE_SHARE_WRITE,
      'share-read+write', LAccessMode, LDetails) or LSharingViolation;
    ProbeShareMode(FILE_SHARE_READ or FILE_SHARE_WRITE or FILE_SHARE_DELETE,
      'share-read+write+delete', LAccessMode, LDetails);

    if LSharingViolation then
    begin
      AErrorDescription := 'Sharing violation under one or more share modes (file likely open by another process):'
        + LDetails;
      Result := True;
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('DiagnoseFileShareModes', AFileName, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.HasReadOnlyAttribute(const APath: string; var AErrorDescription: string): Boolean;
var
  LAttributes: DWORD;
  LErr: DWORD;
begin
  Result := False;
  AErrorDescription := '';

  try
    LAttributes := GetFileAttributes(PChar(ToLongPath(APath)));

    if LAttributes = INVALID_FILE_ATTRIBUTES then
    begin
      LErr := GetLastError;
      AErrorDescription := Format('GetFileAttributes failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

      Exit;
    end;

    if (LAttributes and FILE_ATTRIBUTE_READONLY) <> 0 then
    begin
      // On directories Windows treats this as "special folder" rather than a true write-block,
      // but applications and installers still routinely refuse to write to such folders.
      if (LAttributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
        AErrorDescription := 'Directory has read-only attribute (cosmetic on Win32 but many apps refuse to write here)'
      else
        AErrorDescription := 'File has read-only attribute — write opens via CreateFile will fail with ACCESS_DENIED';

      Result := True;
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('HasReadOnlyAttribute', APath, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.IsEFSEncrypted(const APath: string; var AErrorDescription: string): Boolean;
var
  LAttributes: DWORD;
  LErr: DWORD;
begin
  Result := False;
  AErrorDescription := '';

  try
    LAttributes := GetFileAttributes(PChar(ToLongPath(APath)));

    if LAttributes = INVALID_FILE_ATTRIBUTES then
    begin
      LErr := GetLastError;
      AErrorDescription := Format('GetFileAttributes failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

      Exit;
    end;

    if (LAttributes and FILE_ATTRIBUTE_ENCRYPTED) <> 0 then
    begin
      AErrorDescription := 'EFS-encrypted — only the encrypting user (and designated recovery agents) can decrypt; '
        + 'a different user account, even Administrator, may get ACCESS_DENIED on read';
      Result := True;
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('IsEFSEncrypted', APath, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.IsOnNetworkShare(const APath: string; var AErrorDescription: string): Boolean;
var
  LExpanded: string;
  LDriveType: UINT;
  LRoot: string;
begin
  Result := False;
  AErrorDescription := '';

  try
    LExpanded := ExpandFileName(APath);

    // UNC path: anything starting with \\ (but not the \\?\ long-path or \\.\ device prefix)
    if LExpanded.StartsWith('\\') and not LExpanded.StartsWith('\\?\') and not LExpanded.StartsWith('\\.\') then
    begin
      AErrorDescription := Format('UNC path on network share — SMB share permissions apply in addition to NTFS ACLs: %s',
        [LExpanded]);
      Result := True;
      Exit;
    end;

    // Mapped drive: \\?\ paths or drive letters; GetDriveType wants a root like "X:\".
    if (Length(LExpanded) >= 3) and (LExpanded[2] = ':') then
    begin
      LRoot := LExpanded.Substring(0, 2) + '\';
      LDriveType := GetDriveType(PChar(LRoot));

      if LDriveType = DRIVE_REMOTE then
      begin
        AErrorDescription := Format('Mapped network drive %s — SMB share permissions apply in addition to NTFS ACLs',
          [LRoot]);
        Result := True;
      end;
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('IsOnNetworkShare', APath, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.CurrentUserIsOwner(const APath: string; var AErrorDescription: string): Boolean;
var
  LBytesNeeded: DWORD;
  LSecDesc: PSECURITY_DESCRIPTOR;
  LOwnerSid: PSID;
  LOwnerDefaulted: BOOL;
  LErr: DWORD;
  LSIDName: array[0..255] of Char;
  LDomainName: array[0..255] of Char;
  LSIDNameLen, LDomainNameLen: DWORD;
  LSIDNameUse: SID_NAME_USE;
begin
  Result := True; // optimistic default — if we can't tell, don't fire a false positive
  AErrorDescription := '';
  LBytesNeeded := 0;

  try
    GetFileSecurity(PChar(ToLongPath(APath)), OWNER_SECURITY_INFORMATION, nil, 0, LBytesNeeded);

    if LBytesNeeded = 0 then
    begin
      LErr := GetLastError;
      AErrorDescription := Format('GetFileSecurity(OWNER) failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

      Exit;
    end;

    LSecDesc := AllocMem(NativeInt(LBytesNeeded));
    try
      if not GetFileSecurity(PChar(ToLongPath(APath)), OWNER_SECURITY_INFORMATION, LSecDesc, LBytesNeeded, LBytesNeeded) then
      begin
        LErr := GetLastError;
        AErrorDescription := Format('GetFileSecurity(OWNER) failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

        Exit;
      end;

      if not GetSecurityDescriptorOwner(LSecDesc, LOwnerSid, LOwnerDefaulted) then
      begin
        LErr := GetLastError;
        AErrorDescription := Format('GetSecurityDescriptorOwner failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

        Exit;
      end;

      if LOwnerSid = nil then
        Exit;

      Result := CurrentTokenIsMemberOf(LOwnerSid);

      if not Result then
      begin
        LSIDNameLen := Length(LSIDName);
        LDomainNameLen := Length(LDomainName);

        if LookupAccountSid(nil, LOwnerSid, LSIDName, LSIDNameLen, LDomainName, LDomainNameLen, LSIDNameUse) then
          AErrorDescription := Format('Owner is %s\%s (not current user). Operations that need WRITE_DAC or '
            + 'take-ownership rights will fail unless the owner is in the current token.', [LDomainName, LSIDName])
        else
          AErrorDescription := 'Owner SID does not match current token and could not be resolved to a name';
      end;
    finally
      FreeMem(LSecDesc);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('CurrentUserIsOwner', APath, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.GetEffectiveRightsShortfall(const APath: string; const ACheckWriteRights: Boolean;
  var AErrorDescription: string): Boolean;
type
  PLocalTokenUser = ^TOKEN_USER;
var
  LBytesNeeded: DWORD;
  LSecDesc: PSECURITY_DESCRIPTOR;
  LDACL: PACL;
  LDACLPresent: BOOL;
  LDefaulted: BOOL;
  LErr: DWORD;
  LTokenHandle: THandle;
  LUserBuf: array[0..511] of Byte;
  LRetLen: DWORD;
  LTokenUser: PLocalTokenUser;
  LTrustee: TRUSTEE_W;
  LRights: ACCESS_MASK;
  LRequired: ACCESS_MASK;
  LMissing: ACCESS_MASK;
  LParts: string;
begin
  Result := False;
  AErrorDescription := '';
  LBytesNeeded := 0;
  LRetLen := 0;

  try
    if not OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, LTokenHandle) then
    begin
      LErr := GetLastError;
      AErrorDescription := Format('OpenProcessToken failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

      Exit;
    end;

    try
      if not GetTokenInformation(LTokenHandle, TokenUser, @LUserBuf[0], Length(LUserBuf), LRetLen) then
      begin
        LErr := GetLastError;
        AErrorDescription := Format('GetTokenInformation(TokenUser) failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

        Exit;
      end;
    finally
      CloseHandle(LTokenHandle);
    end;

    LTokenUser := PLocalTokenUser(@LUserBuf[0]);

    GetFileSecurity(PChar(ToLongPath(APath)), DACL_SECURITY_INFORMATION or OWNER_SECURITY_INFORMATION or GROUP_SECURITY_INFORMATION,
      nil, 0, LBytesNeeded);

    if LBytesNeeded = 0 then
    begin
      LErr := GetLastError;
      AErrorDescription := Format('GetFileSecurity for effective rights failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

      Exit;
    end;

    LSecDesc := AllocMem(NativeInt(LBytesNeeded));
    try
      if not GetFileSecurity(PChar(ToLongPath(APath)),
           DACL_SECURITY_INFORMATION or OWNER_SECURITY_INFORMATION or GROUP_SECURITY_INFORMATION,
           LSecDesc, LBytesNeeded, LBytesNeeded) then
      begin
        LErr := GetLastError;
        AErrorDescription := Format('GetFileSecurity for effective rights failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

        Exit;
      end;

      if not GetSecurityDescriptorDacl(LSecDesc, LDACLPresent, LDACL, LDefaulted) then
      begin
        LErr := GetLastError;
        AErrorDescription := Format('GetSecurityDescriptorDacl failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

        Exit;
      end;

      if not LDACLPresent or (LDACL = nil) then
        Exit; // NULL DACL = full access; nothing missing

      FillChar(LTrustee, SizeOf(LTrustee), 0);
      LTrustee.TrusteeForm := TRUSTEE_IS_SID;
      LTrustee.TrusteeType := TRUSTEE_IS_USER;
      // For TRUSTEE_IS_SID the API expects a PSID stuffed into the ptstrName field;
      // both are pointer-sized, so the cast through PWideChar is the documented idiom.
      LTrustee.ptstrName := PWideChar(LTokenUser^.User.Sid);

      LRights := 0;

      if GetEffectiveRightsFromAcl(LDACL^, LTrustee, LRights) <> ERROR_SUCCESS then
      begin
        LErr := GetLastError;
        AErrorDescription := Format('GetEffectiveRightsFromAcl failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

        Exit;
      end;

      // FILE_GENERIC_READ / WRITE / EXECUTE expand to specific bits; checking the
      // ALL_ACCESS subset is too coarse, but the named groups match what a normal
      // application asks for through GENERIC_READ/WRITE.
      if ACheckWriteRights then
        LRequired := FILE_GENERIC_READ or FILE_GENERIC_WRITE
      else
        LRequired := FILE_GENERIC_READ;

      LMissing := LRequired and not LRights;

      if LMissing <> 0 then
      begin
        LParts := '';

        if (LMissing and FILE_READ_DATA) <> 0 then
          LParts := LParts + 'FILE_READ_DATA ';

        if (LMissing and FILE_WRITE_DATA) <> 0 then
          LParts := LParts + 'FILE_WRITE_DATA ';

        if (LMissing and FILE_APPEND_DATA) <> 0 then
          LParts := LParts + 'FILE_APPEND_DATA ';

        if (LMissing and FILE_READ_EA) <> 0 then
          LParts := LParts + 'FILE_READ_EA ';

        if (LMissing and FILE_WRITE_EA) <> 0 then
          LParts := LParts + 'FILE_WRITE_EA ';

        if (LMissing and FILE_READ_ATTRIBUTES) <> 0 then
          LParts := LParts + 'FILE_READ_ATTRIBUTES ';

        if (LMissing and FILE_WRITE_ATTRIBUTES) <> 0 then
          LParts := LParts + 'FILE_WRITE_ATTRIBUTES ';

        if (LMissing and READ_CONTROL) <> 0 then
          LParts := LParts + 'READ_CONTROL ';

        if (LMissing and SYNCHRONIZE) <> 0 then
          LParts := LParts + 'SYNCHRONIZE ';

        AErrorDescription := Format('Effective rights for current user missing: %s(mask: 0x%.8x)', [LParts, LMissing]);

        Result := True;
      end;
    finally
      FreeMem(LSecDesc);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('GetEffectiveRightsShortfall', APath, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.ToLongPath(const ADirectory: string): string;
const
  LONG_PATH_PREFIX     = '\\?\';
  UNC_PREFIX           = '\\';
  LONG_UNC_PATH_PREFIX = '\\?\UNC\';
var
  LNormalized: string;
begin
  if not FOpenFilesLongFileAndPathNameSupport then
    Exit(ADirectory);

  // Already prefixed
  if ADirectory.StartsWith(LONG_PATH_PREFIX) then
    Exit(ADirectory);

  // \\?\ requires an absolute, normalized path — relative paths and '..' segments
  // are not resolved by the kernel when this prefix is used.
  LNormalized := ExpandFileName(ADirectory);

  // UNC path e.g. \\server\share
  if LNormalized.StartsWith(UNC_PREFIX) then
    Result := LONG_UNC_PATH_PREFIX + LNormalized.Substring(2)
  else
    Result := LONG_PATH_PREFIX + LNormalized;
end;

procedure TFileRightsChecker.Execute(const ADirectory: string; const ACheckWriteRights: Boolean);
begin
  var LFiles := TStringList.Create;
  var LSubDirectories := TStringList.Create;
  var LLocalErrorDescription: string := '';

  try
    InitializeDirectoriesAndFiles(ADirectory, LFiles, LSubDirectories, LLocalErrorDescription);

    if LSubDirectories.Count >= 1 then
      DoDirectoryChecks(LSubDirectories, ACheckWriteRights);

    DoFileChecks(LFiles, ACheckWriteRights);
  finally
    LFiles.Free;
    LSubDirectories.Free;
  end;
end;

function TFileRightsChecker.IsDirectoryUnderUACVirtualization(const ADirectory: string; var AErrorDescription: string): Boolean;
var
  LVirtualStorePath: string;
  LSystemDrive: string;
begin
  Result := False;
  AErrorDescription := '';

  try
    // Virtualization only applies to non-elevated 32-bit processes
    if IsRunningElevated(AErrorDescription) then
      Exit;

    if IsProcess64Bit then
      Exit;

    LSystemDrive := GetEnvironmentVariable('SystemDrive');

    // Build the corresponding VirtualStore path
    if not ADirectory.StartsWith(LSystemDrive, True) then
      Exit;

    LVirtualStorePath := GetEnvironmentVariable('LOCALAPPDATA') + '\VirtualStore' + ADirectory.Substring(Length(LSystemDrive));

    // Only report if VirtualStore path actually exists and has content
    if DirectoryExists(LVirtualStorePath) and not IsDirectoryEmpty(LVirtualStorePath) then
    begin
      AErrorDescription := Format('Active UAC virtualization detected — files may be ' + 'redirected to: %s', [LVirtualStorePath]);
      Result := True;
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('IsDirectoryUnderUACVirtualization', ADirectory, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.HasPrivilege(const APrivilegeName: string; var AErrorDescription: string): Boolean;
var
  LTokenHandle: THandle;
  LLUID: TLargeInteger;
  LPrivilegeSet: PRIVILEGE_SET;
  LHasPrivilege: BOOL;
  LErrorCode: DWORD;
begin
  Result := False;
  AErrorDescription := '';

  try
    if not OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, LTokenHandle) then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('OpenProcessToken failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

      Exit;
    end;
    try
      if not LookupPrivilegeValue(nil, PChar(APrivilegeName), LLUID) then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('LookupPrivilegeValue failed for "%s" [%d]: %s', [APrivilegeName, LErrorCode, SysErrorMessage(LErrorCode)]);

        Exit;
      end;

      LPrivilegeSet.PrivilegeCount := 1;
      LPrivilegeSet.Control := PRIVILEGE_SET_ALL_NECESSARY;
      LPrivilegeSet.Privilege[0].Luid := LLUID;
      LPrivilegeSet.Privilege[0].Attributes := 0;

      if not PrivilegeCheck(LTokenHandle, LPrivilegeSet, LHasPrivilege) then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('PrivilegeCheck failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

        Exit;
      end;

      Result := LHasPrivilege;

      if not Result then
        AErrorDescription := Format('Process does not have privilege: %s', [APrivilegeName]);
    finally
      CloseHandle(LTokenHandle);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('HasPrivilege', APrivilegeName, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.IsReparsePoint(const ADirectory: string; var AErrorDescription: string): Boolean;
var
  LAttributes: DWORD;
  LErrorCode: DWORD;
begin
  Result := False;
  AErrorDescription := '';

  try
    LAttributes := GetFileAttributes(PChar(ToLongPath(ADirectory)));

    if LAttributes = INVALID_FILE_ATTRIBUTES then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('GetFileAttributes failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

      Exit;
    end;

    if (LAttributes and FILE_ATTRIBUTE_REPARSE_POINT) <> 0 then
    begin
      AErrorDescription := Format('Path is a reparse point (junction or symlink) — ' + 'target location may have different access rights: %s',
        [ADirectory]);

      Result := True;
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('IsReparsePoint', ADirectory, E, AErrorDescription);
  end;
end;

procedure TFileRightsChecker.GetExceptionErrorDescription(const AErrorMethod, AFileSystemItem: string; const AException: Exception; var AErrorDescription: string);
begin
  AErrorDescription := 'Exception ' + AException.ClassName + ' occurred at ' + AErrorMethod.QuotedString('"')  + ' with message: '
    + AException.Message.QuotedString('"') + '. While checking file system item: ' + AFileSystemItem.QuotedString('"');
end;

function TFileRightsChecker.GetFileIntegrityLevel(const ADirectory: string; var AErrorDescription: string): string;
var
  LFileHandle: THandle;
  LSecDesc: PSECURITY_DESCRIPTOR;
  LBytesNeeded: DWORD;
  LErrorCode: DWORD;
  LLabel: PTOKEN_MANDATORY_LABEL;
  LRIDCount: DWORD;
  LRID: DWORD;
begin
  Result := '';
  AErrorDescription := '';
  LBytesNeeded := 0;

  try
    LFileHandle := CreateFile(PChar(ToLongPath(ADirectory)), READ_CONTROL, FILE_SHARE_READ or FILE_SHARE_WRITE, nil, OPEN_EXISTING,
      FILE_FLAG_BACKUP_SEMANTICS, 0);

    if LFileHandle = INVALID_HANDLE_VALUE then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('CreateFile failed for integrity level check [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

      Exit;
    end;
    try
      // First call to get buffer size
      GetKernelObjectSecurity(LFileHandle, LABEL_SECURITY_INFORMATION, nil, 0, LBytesNeeded);

      if LBytesNeeded = 0 then
      begin
        Result := 'No integrity label — defaults to Medium';

        Exit;
      end;

      LSecDesc := AllocMem(NativeInt(LBytesNeeded));
      try
        if not GetKernelObjectSecurity(LFileHandle, LABEL_SECURITY_INFORMATION, LSecDesc, LBytesNeeded, LBytesNeeded) then
        begin
          LErrorCode := GetLastError;
          AErrorDescription := Format('GetKernelObjectSecurity failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

          Exit;
        end;

        LLabel := PTOKEN_MANDATORY_LABEL(LSecDesc);

        // Get the last sub-authority which is the integrity RID
        LRIDCount := GetSidSubAuthorityCount(LLabel^.Label_.Sid)^;
        LRID := GetSidSubAuthority(LLabel^.Label_.Sid, LRIDCount - 1)^;

        case LRID of
          $0000: Result := 'Untrusted';
          $1000: Result := 'Low';
          $2000: Result := 'Medium';
          $2100: Result := 'Medium Plus';
          $3000: Result := 'High';
          $4000: Result := 'System';
          $5000: Result := 'Protected Process';
          else
            Result := Format('Unknown (RID: 0x%.4x)', [LRID]);
        end;

      finally
        FreeMem(LSecDesc);
      end;
    finally
      CloseHandle(LFileHandle);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('GetFileIntegrityLevel', ADirectory, E, AErrorDescription);
  end;
end;

procedure TFileRightsChecker.CheckProcessBackupPrivileges(const ADirectory: string);
begin
  var LPrivilegeDescription: string := '';

  if not HasPrivilege(SE_BACKUP_NAME, LPrivilegeDescription) then
    LogError(ADirectory, frcMissingPrivilege, LPrivilegeDescription);

  if not HasPrivilege(SE_RESTORE_NAME, LPrivilegeDescription) then
    LogError(ADirectory, frcMissingPrivilege, LPrivilegeDescription);

  if not HasPrivilege(SE_SECURITY_NAME, LPrivilegeDescription) then
    LogError(ADirectory, frcMissingPrivilege, LPrivilegeDescription);
end;

procedure TFileRightsChecker.CheckToAddMoreInfoForCreateFileFailure(const AErrorCode: DWORD; var AErrorDescription: string);
begin
  if AErrorCode = ERROR_ACCESS_DENIED then
  begin
    if IsRunningElevated(AErrorDescription) then
      AErrorDescription := AErrorDescription + ' [Process IS elevated — likely explicit DENY ACE or EFS encryption]'
    else
      AErrorDescription := AErrorDescription
        + ' [Process is NOT elevated — re-run as Administrator to confirm. If that succeeds, likely cause is UAC token filtering,'
        + ' explicit DENY ACE, or EFS encryption]';
  end;
end;

constructor TFileRightsChecker.Create(const AOpenFilesLongFileAndPathNameSupport: Boolean = True; const ACheckProcessBackupPrivileges: Boolean = False;
  const ARunDirectoryGetEffectiveRightsShortfallTests: Boolean = False; const ARunFileGetEffectiveRightsShortfallTests: Boolean = False;
  const ARunCurrentUserIsOwnerTests: Boolean = False);
begin
  inherited Create;

  FReadOnlyStatistics := TStatistics.Create;
  FReadWriteStatistics := TStatistics.Create;
  FErrors := TErrorItemCollection.Create;
  FOpenFilesLongFileAndPathNameSupport := AOpenFilesLongFileAndPathNameSupport;
  FCheckProcessBackupPrivileges := ACheckProcessBackupPrivileges;
  FRunDirectoryGetEffectiveRightsShortfallTests := ARunDirectoryGetEffectiveRightsShortfallTests;
  FRunFileGetEffectiveRightsShortfallTests := ARunFileGetEffectiveRightsShortfallTests;
  FRunCurrentUserIsOwnerTests := ARunCurrentUserIsOwnerTests;
end;

destructor TFileRightsChecker.Destroy;
begin
  FErrors.Free;
  FReadOnlyStatistics.Free;
  FReadWriteStatistics.Free;

  inherited Destroy;
end;

procedure TFileRightsChecker.DoDirectoryChecks(const ADirectories: TStringList; const ACheckWriteRights: Boolean);
begin
  for var LSubDirectory in ADirectories do
  begin
    // Each diagnostic is independent and logs its own error type. Fine-grained on purpose:
    // a single directory can be on a network share, read-only, AND have an explicit deny —
    // we want all three reported, not just the first match.

    var LDesc: string := '';

    if IsOnNetworkShare(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcNetworkShare, LDesc);

    LDesc := '';
    if IsReparsePoint(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcIsReparsePoint, LDesc);

    LDesc := '';
    if HasReadOnlyAttribute(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcReadOnlyAttribute, LDesc);

    LDesc := '';
    if IsEFSEncrypted(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcEFSEncrypted, LDesc);

    LDesc := '';
    if HasEmptyDACL(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcEmptyDACL, LDesc);

    LDesc := '';
    if HasExplicitDenyACE(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcExplicitDenyACE, LDesc);

    LDesc := '';
    if FRunCurrentUserIsOwnerTests then
      if not CurrentUserIsOwner(LSubDirectory, LDesc) and not LDesc.IsEmpty then
        LogError(LSubDirectory, frcOwnershipMismatch, LDesc);

    LDesc := '';
    if FRunDirectoryGetEffectiveRightsShortfallTests then
      if GetEffectiveRightsShortfall(LSubDirectory, ACheckWriteRights, LDesc) then
        LogError(LSubDirectory, frcEffectiveRightsMissing, LDesc);

    LDesc := '';
    if IsDirectoryUnderUACVirtualization(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcUACVirtualization, LDesc);

    // Primary access probe — drives the readable/writable statistic.
    LDesc := '';
    if ACheckWriteRights then
    begin
      if not TestDirectoryWriteRights(LSubDirectory, LDesc) then
        LogError(LSubDirectory, frcDirectoryNotWritable, LDesc)
      else
        FReadWriteStatistics.AddCheckedDirectory;
    end
    else
    begin
      if not TestDirectoryReadRights(LSubDirectory, LDesc) then
        LogError(LSubDirectory, frcDirectoryNotReadable, LDesc)
      else
        FReadOnlyStatistics.AddCheckedDirectory;
    end;
  end;
end;

procedure TFileRightsChecker.DoFileChecks(const AFiles: TStringList; const ACheckWriteRights: Boolean);
begin
  for var LCurrentFile in AFiles do
  begin
    // Same approach as DoDirectoryChecks: each diagnostic gets its own LogError so the
    // operator sees every cause that applies, not just the first.

    var LDesc: string := '';

    if IsOnNetworkShare(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcNetworkShare, LDesc);

    LDesc := '';
    if IsReparsePoint(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcIsReparsePoint, LDesc);

    LDesc := '';
    if HasReadOnlyAttribute(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcReadOnlyAttribute, LDesc);

    LDesc := '';
    if IsEFSEncrypted(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcEFSEncrypted, LDesc);

    LDesc := '';
    if HasEmptyDACL(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcEmptyDACL, LDesc);

    LDesc := '';
    if HasExplicitDenyACE(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcExplicitDenyACE, LDesc);

    LDesc := '';
    if FRunCurrentUserIsOwnerTests then
      if not CurrentUserIsOwner(LCurrentFile, LDesc) and not LDesc.IsEmpty then
        LogError(LCurrentFile, frcOwnershipMismatch, LDesc);

    LDesc := '';
    if FRunFileGetEffectiveRightsShortfallTests then
      if GetEffectiveRightsShortfall(LCurrentFile, ACheckWriteRights, LDesc) then
        LogError(LCurrentFile, frcEffectiveRightsMissing, LDesc);

    LDesc := '';
    if DiagnoseFileShareModes(LCurrentFile, ACheckWriteRights, LDesc) then
      LogError(LCurrentFile, frcShareModeConflict, LDesc);

    // Primary access probe — drives the readable/writable statistic and triggers
    // the "additional info" follow-ups below if it fails.
    LDesc := '';
    if not TestOpenFileRights(LCurrentFile, ACheckWriteRights, LDesc) then
    begin
      var LIntegrityErr: string;
      var LIntegrity: string := GetFileIntegrityLevel(LCurrentFile, LIntegrityErr);
      if LIntegrity <> '' then
        LDesc := LDesc + Format(' | Integrity level: %s', [LIntegrity]) +
          IfThen(LIntegrityErr.IsEmpty, '', ' - ' + LIntegrityErr);

      if ACheckWriteRights then
        LogError(LCurrentFile, frcFileNotWritable, LDesc)
      else
        LogError(LCurrentFile, frcFileNotReadable, LDesc);

      var LExtension := ExtractFileExt(LCurrentFile);
      if MatchText(LExtension, ['.exe', '.dll']) then
      begin
        var LExecDesc: string := '';
        if not TestExecuteRights(LCurrentFile, LExecDesc) then
          LogError(LCurrentFile, frcUserHasNoExecuteRightsForFile, LExecDesc);
      end;
    end
    else if ACheckWriteRights then
      FReadWriteStatistics.AddCheckedFile
    else
      FReadOnlyStatistics.AddCheckedFile;
  end;
end;

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
