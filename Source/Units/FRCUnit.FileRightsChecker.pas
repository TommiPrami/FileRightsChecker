unit FRCUnit.FileRightsChecker;

interface

uses
  Winapi.Windows, System.Classes, System.Generics.Collections, System.SysUtils,
  FRCUnit.WinAPI, FRCUnit.Statistics;

type
  // There is no clean RTL/WinAPI "is this a file or a directory" enum — Win32 only has
  // the FILE_ATTRIBUTE_DIRECTORY bit and the RTL exposes it as the faDirectory flag.
  // A two-value enum is clearer at call sites than a flag check.
  TFileSystemType = (fstDirectory, fstFile);

  // Fired once when the checker starts processing a file or directory.
  TFileSystemItemCallBack = procedure(const AType: TFileSystemType; const AName: string) of object;

  // Fired after each individual test (each diagnostic call) completes. AErrorsCount is
  // the running total of esError-severity findings (warnings and info lines are not
  // counted); AProgress is 0.0 .. 100.0.
  TTestCallBack = procedure(const AType: TFileSystemType; const AName: string; const ATestCount: Integer; const AErrorsCount: Integer;
    const AProgress: Double) of object;

  // Internal: holds the result of a single PreparePass call. Owns its TStringLists;
  // RunPreparedPasses iterates these without touching disk for enumeration.
  TPreparedPass = class(TObject)
  strict private
    FFiles: TStringList;
    FSubDirectories: TStringList;
    FCheckWriteRights: Boolean;
  public
    constructor Create(const ACheckWriteRights: Boolean);
    destructor Destroy; override;

    property Files: TStringList read FFiles;
    property SubDirectories: TStringList read FSubDirectories;
    property CheckWriteRights: Boolean read FCheckWriteRights;
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
    FOnFileSystemItem: TFileSystemItemCallBack;
    FOnTest: TTestCallBack;
    FTotalTestsPlanned: Integer;
    FTestsExecuted: Integer;
    FPreparedPasses: TObjectList<TPreparedPass>;
    FProcessIntegrityRID: Integer;
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
    function GetFileIntegrityLevel(const ADirectory: string; out AIntegrityRID: Integer; var AErrorDescription: string): string;
    function IsDirectoryEmpty(const ADirectory: string): Boolean;
    function DiagnoseFileShareModes(const AFileName: string; const AInReadWriteMode: Boolean; var AErrorDescription: string): Boolean;
    function TryGetAttributes(const APath: string; out AAttributes: DWORD; var AErrorDescription: string): Boolean;
    function IsReparsePoint(const AAttributes: DWORD; const APath: string; var AErrorDescription: string): Boolean;
    function HasReadOnlyAttribute(const AAttributes: DWORD; var AErrorDescription: string): Boolean;
    function IsEFSEncrypted(const AAttributes: DWORD; var AErrorDescription: string): Boolean;
    function HasCloudOrOfflineAttribute(const AAttributes: DWORD; var AErrorDescription: string): Boolean;
    function IsPathTooLong(const APath: string; var AErrorDescription: string): Boolean;
    function IsOnNetworkShare(const APath: string; var AErrorDescription: string): Boolean;
    function CurrentUserIsOwner(const APath: string; var AErrorDescription: string): Boolean;
    function GetEffectiveRightsShortfall(const APath: string; const ACheckWriteRights: Boolean; var AErrorDescription: string): Boolean;
    procedure CheckToAddMoreInfoForCreateFileFailure(const AErrorCode: DWORD; var AErrorDescription: string);
    procedure GetExceptionErrorDescription(const AErrorMethod, AFileSystemItem: string; const AException: Exception; var AErrorDescription: string);
    procedure GetFilesAndDirs(const ADirectory: string; const AFiles, ADirectories: TStringList;  var AErrorDescription: string;
      const AClearLists: Boolean = True);
    procedure LogError(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string;
      const ASeverity: TErrorSeverity = esError);
    procedure CheckProcessBackupPrivileges(const ADirectory: string);
    procedure CheckVolumeInfo(const ADirectory: string; const ACheckWriteRights: Boolean);
    procedure CheckControlledFolderAccess;
    procedure InitializeDirectoriesAndFiles(const ADirectory: string; const AFiles, ASubDirectories: TStringList;
      const ACheckWriteRights: Boolean; var AErrorDescription: string);
    procedure DoDirectoryChecks(const ADirectories: TStringList; const ACheckWriteRights: Boolean);
    procedure DoFileChecks(const AFiles: TStringList; const ACheckWriteRights: Boolean);
    function PerDirectoryTestCount: Integer;
    function PerFileTestCount: Integer;
    procedure ReportItem(const AType: TFileSystemType; const AName: string);
    procedure ReportTest(const AType: TFileSystemType; const AName: string);
  public
    constructor Create(const AOpenFilesLongFileAndPathNameSupport: Boolean = True; const ACheckProcessBackupPrivileges: Boolean = False;
      const ARunDirectoryGetEffectiveRightsShortfallTests: Boolean = False; const ARunFileGetEffectiveRightsShortfallTests: Boolean = False;
      const ARunCurrentUserIsOwnerTests: Boolean = False);
    destructor Destroy; override;

    procedure Execute(const ADirectory: string; const ACheckWriteRights: Boolean);

    // Cumulative-progress API: call PreparePass once per (path, write-mode) pass,
    // then RunPreparedPasses. The total test count is summed across all passes so the
    // progress callback emits a single continuous 0..100% sweep with running test
    // and error totals.
    procedure PreparePass(const ADirectory: string; const ACheckWriteRights: Boolean);
    procedure RunPreparedPasses;

    // One-line summary of who is running the scan and how — user, elevation,
    // process integrity level, bitness, long-path mode. Log this at the top of a
    // run: support cannot interpret customer results without it.
    function RunContextDescription: string;

    property Errors: TErrorItemCollection read FErrors;
    property ReadWriteStatistics: TStatistics read FReadWriteStatistics;
    property ReadOnlyStatistics: TStatistics read FReadOnlyStatistics;
    property RunDirectoryGetEffectiveRightsShortfallTests: Boolean read FRunDirectoryGetEffectiveRightsShortfallTests write FRunDirectoryGetEffectiveRightsShortfallTests;
    property RunFileGetEffectiveRightsShortfallTests: Boolean read FRunFileGetEffectiveRightsShortfallTests write FRunFileGetEffectiveRightsShortfallTests;
    property RunCurrentUserIsOwnerTests: Boolean read FRunCurrentUserIsOwnerTests write FRunCurrentUserIsOwnerTests;
    property OnFileSystemItem: TFileSystemItemCallBack read FOnFileSystemItem write FOnFileSystemItem;
    property OnTest: TTestCallBack read FOnTest write FOnTest;
  end;

implementation

uses
  Winapi.AccCtrl, System.Math, System.StrUtils, System.Win.Registry;

function IsProcess64Bit: Boolean;
begin
  Result := SizeOf(Pointer) = 8;
end;

// Maps a mandatory-integrity RID to its human name. -1 means "could not determine".
function IntegrityRIDToDescription(const ARID: Integer): string;
begin
  case ARID of
    -1:    Result := 'Unknown';
    $0000: Result := 'Untrusted';
    $1000: Result := 'Low';
    $2000: Result := 'Medium';
    $2100: Result := 'Medium Plus';
    $3000: Result := 'High';
    $4000: Result := 'System';
    $5000: Result := 'Protected Process';
    else
      Result := Format('Unknown (RID: 0x%.4x)', [ARID]);
  end;
end;

// Integrity level of the current process token, as the raw RID ($2000 = Medium,
// $3000 = High/elevated...). Returns -1 if it cannot be determined.
function GetProcessIntegrityRID: Integer;
type
  PLocalTokenMandatoryLabel = ^TOKEN_MANDATORY_LABEL;
var
  LTokenHandle: THandle;
  LBuf: array[0..255] of Byte;
  LRetLen: DWORD;
  LLabel: PLocalTokenMandatoryLabel;
  LCount: DWORD;
begin
  Result := -1;
  LRetLen := 0;

  if not OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, LTokenHandle) then
    Exit;

  try
    if not GetTokenInformation(LTokenHandle, TTokenInformationClass(TOKEN_INTEGRITY_LEVEL_INFO_CLASS),
       @LBuf[0], Length(LBuf), LRetLen) then
      Exit;

    LLabel := PLocalTokenMandatoryLabel(@LBuf[0]);
    LCount := GetSidSubAuthorityCount(LLabel^.Label_.Sid)^;
    Result := Integer(GetSidSubAuthority(LLabel^.Label_.Sid, LCount - 1)^);
  finally
    CloseHandle(LTokenHandle);
  end;
end;

// DOMAIN\user of the current process token, or 'Unknown' if it cannot be resolved.
function GetTokenUserName: string;
type
  PLocalTokenUser = ^TOKEN_USER;
var
  LTokenHandle: THandle;
  LBuf: array[0..511] of Byte;
  LRetLen: DWORD;
  LName: array[0..255] of Char;
  LDomain: array[0..255] of Char;
  LNameLen, LDomainLen: DWORD;
  LUse: SID_NAME_USE;
begin
  Result := 'Unknown';
  LRetLen := 0;

  if not OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, LTokenHandle) then
    Exit;

  try
    if not GetTokenInformation(LTokenHandle, TokenUser, @LBuf[0], Length(LBuf), LRetLen) then
      Exit;

    LNameLen := Length(LName);
    LDomainLen := Length(LDomain);

    if LookupAccountSid(nil, PLocalTokenUser(@LBuf[0])^.User.Sid, LName, LNameLen, LDomain, LDomainLen, LUse) then
      Result := string(LDomain) + '\' + string(LName);
  finally
    CloseHandle(LTokenHandle);
  end;
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
  const ACheckWriteRights: Boolean; var AErrorDescription: string);
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

    // Volume-level context (network share, filesystem type, read-only volume,
    // free space) is per-root, so check it once here instead of per item.
    CheckVolumeInfo(LTrimmed, ACheckWriteRights);

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

procedure TFileRightsChecker.LogError(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string;
  const ASeverity: TErrorSeverity = esError);
begin
  var LErrorItem := TErrorItem.Create(AFileSystemItem, AErrorType, AErrorDescription, ASeverity);

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
        CheckToAddMoreInfoForCreateFileFailure(LErrorCode, AErrorDescription);
      end;
    finally
      CloseHandle(LFileHandle);
    end;

    // Write failed: still try to remove the test file so we don't litter the
    // customer's directory; the write error above is what gets reported.
    if not AErrorDescription.IsEmpty then
    begin
      DeleteFile(PChar(ToLongPath(LTempFile)));
      Exit;
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

      CheckToAddMoreInfoForCreateFileFailure(LErrorCode, AErrorDescription);

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

        CheckToAddMoreInfoForCreateFileFailure(LErrorCode, AErrorDescription);

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

      CheckToAddMoreInfoForCreateFileFailure(LErrorCode, AErrorDescription);

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

// Fetches the attribute bits once per item; the individual attribute checks below are
// pure bit tests on this value. One syscall instead of three-plus — matters a lot on
// network shares where every call is a round trip.
function TFileRightsChecker.TryGetAttributes(const APath: string; out AAttributes: DWORD; var AErrorDescription: string): Boolean;
var
  LErr: DWORD;
begin
  Result := False;
  AErrorDescription := '';
  AAttributes := INVALID_FILE_ATTRIBUTES;

  try
    AAttributes := GetFileAttributes(PChar(ToLongPath(APath)));

    if AAttributes = INVALID_FILE_ATTRIBUTES then
    begin
      LErr := GetLastError;
      AErrorDescription := Format('GetFileAttributes failed [%d]: %s', [LErr, SysErrorMessage(LErr)]);

      CheckToAddMoreInfoForCreateFileFailure(LErr, AErrorDescription);

      Exit;
    end;

    Result := True;
  except
    on E: Exception do
      GetExceptionErrorDescription('TryGetAttributes', APath, E, AErrorDescription);
  end;
end;

function TFileRightsChecker.IsReparsePoint(const AAttributes: DWORD; const APath: string; var AErrorDescription: string): Boolean;
begin
  Result := (AAttributes and FILE_ATTRIBUTE_REPARSE_POINT) <> 0;

  if Result then
    AErrorDescription := Format('Path is a reparse point (junction or symlink) — target location may have different access rights: %s',
      [APath])
  else
    AErrorDescription := '';
end;

function TFileRightsChecker.HasReadOnlyAttribute(const AAttributes: DWORD; var AErrorDescription: string): Boolean;
begin
  Result := (AAttributes and FILE_ATTRIBUTE_READONLY) <> 0;

  if Result then
  begin
    // On directories Windows treats this as "special folder" rather than a true write-block,
    // but applications and installers still routinely refuse to write to such folders.
    if (AAttributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
      AErrorDescription := 'Directory has read-only attribute (cosmetic on Win32 but many apps refuse to write here)'
    else
      AErrorDescription := 'File has read-only attribute — write opens via CreateFile will fail with ACCESS_DENIED';
  end
  else
    AErrorDescription := '';
end;

function TFileRightsChecker.IsEFSEncrypted(const AAttributes: DWORD; var AErrorDescription: string): Boolean;
begin
  Result := (AAttributes and FILE_ATTRIBUTE_ENCRYPTED) <> 0;

  if Result then
    AErrorDescription := 'EFS-encrypted — only the encrypting user (and designated recovery agents) can decrypt; '
      + 'a different user account, even Administrator, may get ACCESS_DENIED on read'
  else
    AErrorDescription := '';
end;

function TFileRightsChecker.HasCloudOrOfflineAttribute(const AAttributes: DWORD; var AErrorDescription: string): Boolean;
begin
  Result := (AAttributes and (FILE_ATTRIBUTE_OFFLINE_BIT or FILE_ATTRIBUTE_RECALL_ON_OPEN
    or FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS)) <> 0;

  if Result then
    AErrorDescription := 'File is offline or a cloud placeholder (OneDrive, Dropbox, HSM) — opening it triggers '
      + 'network hydration and can hang, fail, or succeed only when the sync client is running and signed in'
  else
    AErrorDescription := '';
end;

function TFileRightsChecker.IsPathTooLong(const APath: string; var AErrorDescription: string): Boolean;
begin
  Result := Length(APath) >= MAX_PATH;

  if Result then
  begin
    AErrorDescription := Format('Full path is %d characters (MAX_PATH is 260) — applications built without '
      + 'long-path support cannot open this item even with correct permissions', [Length(APath)]);

    if not FOpenFilesLongFileAndPathNameSupport then
      AErrorDescription := AErrorDescription
        + ' [long-path support is OFF in this scan too, so the probes below may fail for this same reason]';
  end
  else
    AErrorDescription := '';
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

      // GetEffectiveRightsFromAcl returns its error code directly — it does NOT
      // set GetLastError like BOOL-returning APIs do.
      LErr := GetEffectiveRightsFromAcl(LDACL^, LTrustee, LRights);

      if LErr <> ERROR_SUCCESS then
      begin
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
  // Single-pass convenience wrapper: clear any queued passes, prepare this one, run.
  // Use PreparePass + RunPreparedPasses directly when you need cumulative progress
  // across multiple passes (e.g. RW + RO in one continuous sweep).
  FPreparedPasses.Clear;
  PreparePass(ADirectory, ACheckWriteRights);
  RunPreparedPasses;
end;

procedure TFileRightsChecker.PreparePass(const ADirectory: string; const ACheckWriteRights: Boolean);
begin
  // Starting a fresh batch — first PreparePass after a RunPreparedPasses (or after
  // construction) zeroes the progress counters before accumulating.
  if FPreparedPasses.Count = 0 then
  begin
    FTotalTestsPlanned := 0;
    FTestsExecuted := 0;
  end;

  var LPass := TPreparedPass.Create(ACheckWriteRights);
  try
    var LLocalErrorDescription: string := '';
    InitializeDirectoriesAndFiles(ADirectory, LPass.Files, LPass.SubDirectories, ACheckWriteRights, LLocalErrorDescription);

    Inc(FTotalTestsPlanned,
      LPass.SubDirectories.Count * PerDirectoryTestCount + LPass.Files.Count * PerFileTestCount);

    FPreparedPasses.Add(LPass);
    LPass := nil; // ownership transferred
  finally
    LPass.Free; // no-op if transferred
  end;
end;

procedure TFileRightsChecker.RunPreparedPasses;
begin
  try
    // Defender Controlled Folder Access is per-application, not per-path, so one
    // check covers the whole run. Only meaningful when writes are being tested.
    for var LPass in FPreparedPasses do
      if LPass.CheckWriteRights then
      begin
        CheckControlledFolderAccess;
        Break;
      end;

    for var LPass in FPreparedPasses do
    begin
      if LPass.SubDirectories.Count >= 1 then
        DoDirectoryChecks(LPass.SubDirectories, LPass.CheckWriteRights);

      DoFileChecks(LPass.Files, LPass.CheckWriteRights);
    end;

    // Guarantee a final 100% notification across the whole batch even if rounding
    // or skipped sub-tests left the counter a hair short.
    if (FTotalTestsPlanned > 0) and (FTestsExecuted < FTotalTestsPlanned) then
    begin
      FTestsExecuted := FTotalTestsPlanned;
      if Assigned(FOnTest) then
        FOnTest(fstFile, '', FTestsExecuted, FErrors.ErrorCount, 100.0);
    end;
  finally
    // Drop the prepared passes so the next PreparePass call starts a fresh batch.
    // Counters stay as-is so callers can read final state right after RunPreparedPasses.
    FPreparedPasses.Clear;
  end;
end;

function TFileRightsChecker.RunContextDescription: string;
var
  LElevationError: string;
begin
  Result := Format('Run context: User: %s | Elevated: %s | Process integrity: %s | Process: %s | Long path support: %s',
    [GetTokenUserName,
     IfThen(IsRunningElevated(LElevationError), 'Yes', 'No'),
     IntegrityRIDToDescription(FProcessIntegrityRID),
     IfThen(IsProcess64Bit, '64-bit', '32-bit'),
     IfThen(FOpenFilesLongFileAndPathNameSupport, 'On', 'Off')]);
end;

function TFileRightsChecker.PerDirectoryTestCount: Integer;
begin
  // Must match the number of ReportTest calls in DoDirectoryChecks.
  // 9 unconditional (4 attribute tests, path length, DACL, deny ACE, UAC
  // virtualization, primary probe) + up to 2 gated.
  Result := 9;

  if FRunCurrentUserIsOwnerTests then
    Inc(Result);

  if FRunDirectoryGetEffectiveRightsShortfallTests then
    Inc(Result);
end;

function TFileRightsChecker.PerFileTestCount: Integer;
begin
  // Must match the number of ReportTest calls in DoFileChecks.
  // 9 unconditional (4 attribute tests, path length, DACL, deny ACE, share modes,
  // primary probe) + up to 2 gated. (TestExecuteRights / GetFileIntegrityLevel are
  // sub-steps of the primary access probe and don't count separately.)
  Result := 9;

  if FRunCurrentUserIsOwnerTests then
    Inc(Result);

  if FRunFileGetEffectiveRightsShortfallTests then
    Inc(Result);
end;

procedure TFileRightsChecker.ReportItem(const AType: TFileSystemType; const AName: string);
begin
  if Assigned(FOnFileSystemItem) then
    FOnFileSystemItem(AType, AName);
end;

procedure TFileRightsChecker.ReportTest(const AType: TFileSystemType; const AName: string);
var
  LProgress: Double;
begin
  Inc(FTestsExecuted);

  if not Assigned(FOnTest) then
    Exit;

  if FTotalTestsPlanned > 0 then
    LProgress := (FTestsExecuted / FTotalTestsPlanned) * 100.0
  else
    LProgress := 0;

  // Defensive clamp — nested or unaccounted tests can push us over briefly.
  if LProgress > 100.0 then
    LProgress := 100.0;

  FOnTest(AType, AName, FTestsExecuted, FErrors.ErrorCount, LProgress);
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

procedure TFileRightsChecker.GetExceptionErrorDescription(const AErrorMethod, AFileSystemItem: string; const AException: Exception; var AErrorDescription: string);
begin
  AErrorDescription := 'Exception ' + AException.ClassName + ' occurred at ' + AErrorMethod.QuotedString('"')  + ' with message: '
    + AException.Message.QuotedString('"') + '. While checking file system item: ' + AFileSystemItem.QuotedString('"');
end;

function TFileRightsChecker.GetFileIntegrityLevel(const ADirectory: string; out AIntegrityRID: Integer;
  var AErrorDescription: string): string;
var
  LFileHandle: THandle;
  LSecDesc: PSECURITY_DESCRIPTOR;
  LBytesNeeded: DWORD;
  LErrorCode: DWORD;
  LLabel: PTOKEN_MANDATORY_LABEL;
  LRIDCount: DWORD;
begin
  Result := '';
  AErrorDescription := '';
  AIntegrityRID := -1;
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
        // Unlabeled objects are implicitly Medium.
        AIntegrityRID := $2000;
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
        AIntegrityRID := Integer(GetSidSubAuthority(LLabel^.Label_.Sid, LRIDCount - 1)^);

        Result := IntegrityRIDToDescription(AIntegrityRID);
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

  // Missing backup/restore/security privileges are normal for non-elevated processes,
  // so these are warnings (context), not errors.
  if not HasPrivilege(SE_BACKUP_NAME, LPrivilegeDescription) then
    LogError(ADirectory, frcMissingPrivilege, LPrivilegeDescription, esWarning);

  if not HasPrivilege(SE_RESTORE_NAME, LPrivilegeDescription) then
    LogError(ADirectory, frcMissingPrivilege, LPrivilegeDescription, esWarning);

  if not HasPrivilege(SE_SECURITY_NAME, LPrivilegeDescription) then
    LogError(ADirectory, frcMissingPrivilege, LPrivilegeDescription, esWarning);
end;

// Volume-level facts that decide read/writeability before any ACL is even looked at.
// Called once per top-level scan root — not per item.
procedure TFileRightsChecker.CheckVolumeInfo(const ADirectory: string; const ACheckWriteRights: Boolean);
const
  LOW_DISK_SPACE_WARNING_BYTES = Int64(100) * 1024 * 1024; // 100 MB
var
  LRootBuf: array[0..MAX_PATH] of Char;
  LFSNameBuf: array[0..63] of Char;
  LFlags: DWORD;
  LMaxComponentLength: DWORD;
  LRoot: string;
  LFreeAvailable: Int64;
  LTotal: Int64;
  LDesc: string;
begin
  try
    // Network share is per-root context — SMB share permissions apply on top of NTFS.
    LDesc := '';
    if IsOnNetworkShare(ADirectory, LDesc) then
      LogError(ADirectory, frcNetworkShare, LDesc, esInfo);

    // GetVolumePathName resolves mount points correctly ('C:\Mount\Data\x' can live
    // on a different volume than C:).
    if GetVolumePathName(PChar(ADirectory), LRootBuf, Length(LRootBuf)) then
      LRoot := LRootBuf
    else
      LRoot := IncludeTrailingPathDelimiter(ExtractFileDrive(ADirectory));

    LFlags := 0;
    LMaxComponentLength := 0;

    if GetVolumeInformation(PChar(LRoot), nil, 0, nil, LMaxComponentLength, LFlags, LFSNameBuf, Length(LFSNameBuf)) then
    begin
      var LFSName: string := LFSNameBuf;

      // FAT32/exFAT have no ACLs at all: every ACL-flavored finding is moot, and
      // per-user permissions can never be the cause of failures on these volumes.
      if not SameText(LFSName, 'NTFS') and not SameText(LFSName, 'ReFS') then
        LogError(ADirectory, frcNonNTFSFileSystem,
          Format('File system on %s is %s — it has no NTFS ACLs, so per-user permissions cannot be the cause of access '
            + 'failures here (and ACL-related findings do not apply)', [LRoot, LFSName]), esInfo);

      if ACheckWriteRights and ((LFlags and FILE_READ_ONLY_VOLUME) <> 0) then
        LogError(ADirectory, frcReadOnlyVolume,
          Format('Volume %s is READ-ONLY — every write fails regardless of ACLs or elevation', [LRoot]));
    end;

    if ACheckWriteRights and GetDiskFreeSpaceEx(PChar(LRoot), LFreeAvailable, LTotal, nil) then
      if LFreeAvailable < LOW_DISK_SPACE_WARNING_BYTES then
        LogError(ADirectory, frcLowDiskSpace,
          Format('Only %d MB free for the current user on %s — writes may fail with disk-full or quota errors',
            [LFreeAvailable div (1024 * 1024), LRoot]), esWarning);
  except
    on E: Exception do
    begin
      var LExceptionDesc: string := '';
      GetExceptionErrorDescription('CheckVolumeInfo', ADirectory, E, LExceptionDesc);
      LogError(ADirectory, frcNone, LExceptionDesc, esWarning);
    end;
  end;
end;

// Windows Defender Controlled Folder Access denies writes per APPLICATION, regardless
// of ACLs or elevation — a classic cause of "ACCESS_DENIED but the permissions look
// fine". Checked once per run (best effort; the key may be unreadable).
procedure TFileRightsChecker.CheckControlledFolderAccess;
const
  CFA_KEY = 'SOFTWARE\Microsoft\Windows Defender\Windows Defender Exploit Guard\Controlled Folder Access';
  CFA_VALUE = 'EnableControlledFolderAccess';
begin
  try
    var LRegistry := TRegistry.Create(KEY_READ or KEY_WOW64_64KEY);
    try
      LRegistry.RootKey := HKEY_LOCAL_MACHINE;

      if LRegistry.OpenKeyReadOnly(CFA_KEY) and LRegistry.ValueExists(CFA_VALUE) then
        case LRegistry.ReadInteger(CFA_VALUE) of
          1: LogError('<system>', frcControlledFolderAccess,
               'Windows Defender Controlled Folder Access is ENABLED (block mode) — writes into protected folders are '
               + 'denied per-application regardless of NTFS ACLs or elevation. If only this application fails with '
               + 'ACCESS_DENIED, check Defender''s protected-folders and allowed-apps lists', esWarning);
          2: LogError('<system>', frcControlledFolderAccess,
               'Windows Defender Controlled Folder Access is in AUDIT mode — writes are allowed but audited', esInfo);
        end;
    finally
      LRegistry.Free;
    end;
  except
    // The Defender registry area can be ACL-restricted; this is best-effort context,
    // so silently skip when unreadable.
  end;
end;

// Appends a plain-language likely cause to a failed CreateFile / GetFileAttributes
// error message, based on the Win32 error code. This is the piece of text a support
// person actually acts on.
procedure TFileRightsChecker.CheckToAddMoreInfoForCreateFileFailure(const AErrorCode: DWORD; var AErrorDescription: string);
var
  LElevationDummy: string;
begin
  case AErrorCode of
    ERROR_ACCESS_DENIED:
      if IsRunningElevated(LElevationDummy) then
        AErrorDescription := AErrorDescription
          + ' [Process IS elevated — likely explicit DENY ACE, EFS encryption, mandatory integrity label,'
          + ' or Defender Controlled Folder Access]'
      else
        AErrorDescription := AErrorDescription
          + ' [Process is NOT elevated — re-run as Administrator to confirm. If that succeeds, likely cause is UAC token filtering,'
          + ' explicit DENY ACE, or EFS encryption]';

    ERROR_WRITE_PROTECT:
      AErrorDescription := AErrorDescription + ' [Media or volume is write-protected — no ACL change will help]';

    ERROR_SHARING_VIOLATION:
      AErrorDescription := AErrorDescription + ' [File is open in another process with an incompatible share mode — not a permissions problem]';

    ERROR_LOCK_VIOLATION:
      AErrorDescription := AErrorDescription + ' [Another process holds a byte-range lock on the file — not a permissions problem]';

    ERROR_HANDLE_DISK_FULL, ERROR_DISK_FULL:
      AErrorDescription := AErrorDescription + ' [Disk is full — writes fail regardless of permissions]';

    ERROR_DISK_QUOTA_EXCEEDED:
      AErrorDescription := AErrorDescription + ' [Per-user disk quota exceeded — an administrator or another user may still be able to write]';

    ERROR_FILENAME_EXCED_RANGE:
      AErrorDescription := AErrorDescription + ' [Path exceeds MAX_PATH (260) — the application needs long-path support to open this]';

    ERROR_NETWORK_ACCESS_DENIED:
      AErrorDescription := AErrorDescription + ' [Denied at the SMB SHARE level — check share permissions on the server, not NTFS ACLs]';

    ERROR_INVALID_NAME:
      AErrorDescription := AErrorDescription + ' [Path contains characters or a form Windows cannot parse — not a permissions problem]';

    ERROR_CANT_ACCESS_FILE:
      AErrorDescription := AErrorDescription + ' [File cannot be accessed by the system — often a broken reparse point or a cloud placeholder whose sync provider is unavailable]';
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
  FPreparedPasses := TObjectList<TPreparedPass>.Create(True);
  // Cached once — needed for the file-vs-process integrity comparison on failures.
  FProcessIntegrityRID := GetProcessIntegrityRID;
  FOpenFilesLongFileAndPathNameSupport := AOpenFilesLongFileAndPathNameSupport;
  FCheckProcessBackupPrivileges := ACheckProcessBackupPrivileges;
  FRunDirectoryGetEffectiveRightsShortfallTests := ARunDirectoryGetEffectiveRightsShortfallTests;
  FRunFileGetEffectiveRightsShortfallTests := ARunFileGetEffectiveRightsShortfallTests;
  FRunCurrentUserIsOwnerTests := ARunCurrentUserIsOwnerTests;
end;

destructor TFileRightsChecker.Destroy;
begin
  FPreparedPasses.Free;
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
    // a single directory can be a reparse point, read-only, AND have an explicit deny —
    // we want all three reported, not just the first match.
    //
    // Each diagnostic is followed by ReportTest so progress accounting stays in lockstep
    // with PerDirectoryTestCount.

    ReportItem(fstDirectory, LSubDirectory);

    // Attributes are fetched ONCE; the four attribute tests below are pure bit tests.
    var LDesc: string := '';
    var LAttributes: DWORD;

    if TryGetAttributes(LSubDirectory, LAttributes, LDesc) then
    begin
      LDesc := '';
      if IsReparsePoint(LAttributes, LSubDirectory, LDesc) then
        LogError(LSubDirectory, frcIsReparsePoint, LDesc, esWarning);
      ReportTest(fstDirectory, LSubDirectory);

      LDesc := '';
      if HasReadOnlyAttribute(LAttributes, LDesc) then
        LogError(LSubDirectory, frcReadOnlyAttribute, LDesc, esInfo);
      ReportTest(fstDirectory, LSubDirectory);

      LDesc := '';
      if IsEFSEncrypted(LAttributes, LDesc) then
        LogError(LSubDirectory, frcEFSEncrypted, LDesc, esWarning);
      ReportTest(fstDirectory, LSubDirectory);

      LDesc := '';
      if HasCloudOrOfflineAttribute(LAttributes, LDesc) then
        LogError(LSubDirectory, frcCloudPlaceholder, LDesc, esWarning);
      ReportTest(fstDirectory, LSubDirectory);
    end
    else
    begin
      LogError(LSubDirectory, frcDirectoryNotReadable, LDesc, esWarning);

      // The four attribute tests cannot run — keep the progress counter in sync.
      for var LSkipped := 1 to 4 do
        ReportTest(fstDirectory, LSubDirectory);
    end;

    LDesc := '';
    if IsPathTooLong(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcPathTooLong, LDesc, esWarning);
    ReportTest(fstDirectory, LSubDirectory);

    LDesc := '';
    if HasEmptyDACL(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcEmptyDACL, LDesc);
    ReportTest(fstDirectory, LSubDirectory);

    LDesc := '';
    if HasExplicitDenyACE(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcExplicitDenyACE, LDesc);
    ReportTest(fstDirectory, LSubDirectory);

    if FRunCurrentUserIsOwnerTests then
    begin
      LDesc := '';
      if not CurrentUserIsOwner(LSubDirectory, LDesc) and not LDesc.IsEmpty then
        LogError(LSubDirectory, frcOwnershipMismatch, LDesc, esInfo);
      ReportTest(fstDirectory, LSubDirectory);
    end;

    if FRunDirectoryGetEffectiveRightsShortfallTests then
    begin
      LDesc := '';
      if GetEffectiveRightsShortfall(LSubDirectory, ACheckWriteRights, LDesc) then
        LogError(LSubDirectory, frcEffectiveRightsMissing, LDesc);
      ReportTest(fstDirectory, LSubDirectory);
    end;

    LDesc := '';
    if IsDirectoryUnderUACVirtualization(LSubDirectory, LDesc) then
      LogError(LSubDirectory, frcUACVirtualization, LDesc, esWarning);
    ReportTest(fstDirectory, LSubDirectory);

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
    ReportTest(fstDirectory, LSubDirectory);
  end;
end;

procedure TFileRightsChecker.DoFileChecks(const AFiles: TStringList; const ACheckWriteRights: Boolean);
begin
  for var LCurrentFile in AFiles do
  begin
    // Same approach as DoDirectoryChecks: each diagnostic gets its own LogError so the
    // operator sees every cause that applies, not just the first.
    //
    // Each ReportTest call must be matched by PerFileTestCount to keep progress in sync.

    ReportItem(fstFile, LCurrentFile);

    // Attributes are fetched ONCE; the four attribute tests below are pure bit tests.
    var LDesc: string := '';
    var LAttributes: DWORD;

    if TryGetAttributes(LCurrentFile, LAttributes, LDesc) then
    begin
      LDesc := '';
      if IsReparsePoint(LAttributes, LCurrentFile, LDesc) then
        LogError(LCurrentFile, frcIsReparsePoint, LDesc, esWarning);
      ReportTest(fstFile, LCurrentFile);

      LDesc := '';
      if HasReadOnlyAttribute(LAttributes, LDesc) then
        LogError(LCurrentFile, frcReadOnlyAttribute, LDesc, esWarning);
      ReportTest(fstFile, LCurrentFile);

      LDesc := '';
      if IsEFSEncrypted(LAttributes, LDesc) then
        LogError(LCurrentFile, frcEFSEncrypted, LDesc, esWarning);
      ReportTest(fstFile, LCurrentFile);

      LDesc := '';
      if HasCloudOrOfflineAttribute(LAttributes, LDesc) then
        LogError(LCurrentFile, frcCloudPlaceholder, LDesc, esWarning);
      ReportTest(fstFile, LCurrentFile);
    end
    else
    begin
      LogError(LCurrentFile, frcFileNotReadable, LDesc, esWarning);

      // The four attribute tests cannot run — keep the progress counter in sync.
      for var LSkipped := 1 to 4 do
        ReportTest(fstFile, LCurrentFile);
    end;

    LDesc := '';
    if IsPathTooLong(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcPathTooLong, LDesc, esWarning);
    ReportTest(fstFile, LCurrentFile);

    LDesc := '';
    if HasEmptyDACL(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcEmptyDACL, LDesc);
    ReportTest(fstFile, LCurrentFile);

    LDesc := '';
    if HasExplicitDenyACE(LCurrentFile, LDesc) then
      LogError(LCurrentFile, frcExplicitDenyACE, LDesc);
    ReportTest(fstFile, LCurrentFile);

    if FRunCurrentUserIsOwnerTests then
    begin
      LDesc := '';
      if not CurrentUserIsOwner(LCurrentFile, LDesc) and not LDesc.IsEmpty then
        LogError(LCurrentFile, frcOwnershipMismatch, LDesc, esInfo);
      ReportTest(fstFile, LCurrentFile);
    end;

    if FRunFileGetEffectiveRightsShortfallTests then
    begin
      LDesc := '';
      if GetEffectiveRightsShortfall(LCurrentFile, ACheckWriteRights, LDesc) then
        LogError(LCurrentFile, frcEffectiveRightsMissing, LDesc);
      ReportTest(fstFile, LCurrentFile);
    end;

    // In-use files are context, not failures: the permissive-share probe in
    // TestOpenFileRights below still decides readable/writable.
    LDesc := '';
    if DiagnoseFileShareModes(LCurrentFile, ACheckWriteRights, LDesc) then
      LogError(LCurrentFile, frcShareModeConflict, LDesc, esInfo);
    ReportTest(fstFile, LCurrentFile);

    // Primary access probe — drives the readable/writable statistic and triggers
    // the "additional info" follow-ups below if it fails.
    LDesc := '';
    if not TestOpenFileRights(LCurrentFile, ACheckWriteRights, LDesc) then
    begin
      var LIntegrityErr: string;
      var LFileIntegrityRID: Integer;
      var LIntegrity: string := GetFileIntegrityLevel(LCurrentFile, LFileIntegrityRID, LIntegrityErr);
      if LIntegrity <> '' then
        LDesc := LDesc + Format(' | Integrity level: %s', [LIntegrity]) +
          IfThen(LIntegrityErr.IsEmpty, '', ' - ' + LIntegrityErr);

      // Mandatory integrity: a file labeled above the process level is write-blocked
      // by the no-write-up policy no matter what the ACLs say. This is exactly the
      // kind of "permissions look fine but writes fail" case the tool exists for.
      if (FProcessIntegrityRID >= 0) and (LFileIntegrityRID > FProcessIntegrityRID) then
        LDesc := LDesc + Format(' | MANDATORY INTEGRITY BLOCK: file integrity (%s) is above process integrity (%s) — '
          + 'write access is denied by the no-write-up policy regardless of ACLs',
          [IntegrityRIDToDescription(LFileIntegrityRID), IntegrityRIDToDescription(FProcessIntegrityRID)]);

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
    ReportTest(fstFile, LCurrentFile);
  end;
end;

{ TPreparedPass }

constructor TPreparedPass.Create(const ACheckWriteRights: Boolean);
begin
  inherited Create;

  FFiles := TStringList.Create;
  FSubDirectories := TStringList.Create;
  FCheckWriteRights := ACheckWriteRights;
end;

destructor TPreparedPass.Destroy;
begin
  FFiles.Free;
  FSubDirectories.Free;

  inherited Destroy;
end;

end.
