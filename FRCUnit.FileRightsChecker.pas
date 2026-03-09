unit FRCUnit.FileRightsChecker;

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Generics.Collections;

const
  ACCESS_ALLOWED_ACE_TYPE         = BYTE($0);
  ACCESS_DENIED_ACE_TYPE          = BYTE($1);
  SYSTEM_AUDIT_ACE_TYPE           = BYTE($2);
  SYSTEM_ALARM_ACE_TYPE           = BYTE($3);
  ACCESS_ALLOWED_COMPOUND_ACE_TYPE = BYTE($4);
  ACCESS_ALLOWED_OBJECT_ACE_TYPE  = BYTE($5);
  ACCESS_DENIED_OBJECT_ACE_TYPE   = BYTE($6);
  SYSTEM_AUDIT_OBJECT_ACE_TYPE    = BYTE($7);

type
  TFileRightErrorType = (frcMissingPrivilege, frcFileNotReadable, frcFileNotWritable, frcUserHasNoExecuteRightsForFile,
    frcIsReparsePoint, frcUACVirtualization, frcDirectoryNotReadable, frcDirectoryNotWritable);

  TACE_HEADER = record
    AceType:  BYTE;
    AceFlags: BYTE;
    AceSize:  WORD;
  end;
  PACE_HEADER = ^TACE_HEADER;

  TACCESS_DENIED_ACE = packed record
    Header:   TACE_HEADER;
    Mask:     ACCESS_MASK;
    SidStart: DWORD;
  end;
  PACCESS_DENIED_ACE = ^TACCESS_DENIED_ACE;

  TACL_SIZE_INFORMATION = record
    AceCount:      DWORD;
    AclBytesInUse: DWORD;
    AclBytesFree:  DWORD;
  end;

  ACL_INFORMATION_CLASS = (AclRevisionInformation = 1, AclSizeInformation = 2);

  TOKEN_MANDATORY_LABEL = record
    Label_: SID_AND_ATTRIBUTES;
  end;
  PTOKEN_MANDATORY_LABEL = ^TOKEN_MANDATORY_LABEL;

  TErrorItem = class(TObject)
  strict private
    FFileSystemItem: string;
    FErrorType: TFileRightErrorType;
    FErrorDescription: string;
    function GetErrorTypeStr: string;
  public
    constructor Create(const AFileSystemItem: string; const AErrorType: TFileRightErrorType;
      const AErrorDescription: string);

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
    function GetTempFileName(const ADirectory: string): string;
    function HasExplicitDenyACE(const ADirectory: string; var AErrorDescription: string): Boolean;
    function TestDirectoryReadRights(const ADirectory: string; var AErrorDescription: string): Boolean;
    function TestDirectoryWriteRights(const ADirectory: string; var AErrorDescription: string): Boolean;
    function TestExecuteRights(const AFileName: string; var AErrorDescription: string): Boolean;
    function TestOpenFileRights(const AFileName: string; const AInReadWriteMode: Boolean; var AErrorDescription: string): Boolean;
    function IsRunningElevated: Boolean;
    function ToLongPath(const ADirectory: string): string;
    function HasEmptyDACL(const ADirectory: string; var AErrorDescription: string): Boolean;
    function IsDirectoryUnderUACVirtualization(const ADirectory: string; var AErrorDescription: string): Boolean;
    function HasPrivilege(const APrivilegeName: string; var AErrorDescription: string): Boolean;
    function IsReparsePoint(const ADirectory: string; var AErrorDescription: string): Boolean;
    function GetFileIntegrityLevel(const ADirectory: string; var AErrorDescription: string): string;
    function IsDirectoryEmpty(const ADirectory: string): Boolean;
    procedure CheckToAddMoreInfoForCreateFileFailure(const AErrorCode: DWORD; var AErrorDescription: string);
    procedure GetExceptionErrorDescription(const AErrorMethod, AFileSystemItem: string; const AException: Exception; var AErrorDescription: string);
    procedure GetFilesAndDirs(const ADirectory: string; const AFiles, ADirectories: TStringList;  const AClearLists: Boolean = True);
    procedure LogError(const AFileSystemItem: string; const AErrorType: TFileRightErrorType; const AErrorDescription: string);
    procedure CheckProcessBackupPrivileges(const ADirectory: string);
    procedure InitializeDirectoriesAndFiles(const ADirectory: string; const AFiles, ASubDirectories: TStringList);
    procedure DoDirectoryChecks(const ADirectories: TStringList; const ACheckWriteRights: Boolean);
    procedure DoFileChecks(const AFiles: TStringList; const ACheckWriteRights: Boolean);
  public
    constructor Create(const AOpenFilesLongFileAndPathNameSupport: Boolean = True; const ACheckProcessBackupPrivileges: Boolean = False);
    destructor Destroy; override;

    procedure Execute(const ADirectory: string; const ACheckWriteRights: Boolean);
    property Errors: TErrorItemCollection read FErrors;
    property ReadWriteStatistics: TStatistics read FReadWriteStatistics;
    property ReadOnlyStatistics: TStatistics read FReadOnlyStatistics;
  end;

implementation

uses
  System.Math, System.StrUtils;

function IsProcess64Bit: Boolean;
begin
  Result := SizeOf(Pointer) = 8;
end;

{ TFileRightsChecker }

procedure TFileRightsChecker.InitializeDirectoriesAndFiles(const ADirectory: string; const AFiles, ASubDirectories: TStringList);
begin
  var LDirectories := ADirectory.Split([';']);

  for var LDirectory in LDirectories do
  begin
    GetFilesAndDirs(LDirectory, AFiles, ASubDirectories, False);
    ASubDirectories.Insert(0, LDirectory);

    if FCheckProcessBackupPrivileges then
      CheckProcessBackupPrivileges(LDirectory);
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
    GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION,
      nil, 0, LBytesNeeded);

    if LBytesNeeded = 0 then
    begin
      AErrorDescription := Format('GetFileSecurity failed [%d]: %s',
        [GetLastError, SysErrorMessage(GetLastError)]);
      Exit;
    end;

    LSecDesc := AllocMem(LBytesNeeded);
    try
      if not GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION,
         LSecDesc, LBytesNeeded, LBytesNeeded) then
      begin
        AErrorDescription := Format('GetFileSecurity failed [%d]: %s',
          [GetLastError, SysErrorMessage(GetLastError)]);
        Exit;
      end;

      if not GetSecurityDescriptorDacl(LSecDesc, LDACLPresent, LDACL, LDefaulted) then
      begin
        AErrorDescription := Format('GetSecurityDescriptorDacl failed [%d]: %s',
          [GetLastError, SysErrorMessage(GetLastError)]);
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

function TFileRightsChecker.IsRunningElevated: Boolean;
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
    begin
      // TODO: Log Error
      raise;
    end;
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
    GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION,
      nil, 0, LBytesNeeded);

    if LBytesNeeded = 0 then
    begin
      AErrorDescription := Format('GetFileSecurity failed [%d]: %s',
        [GetLastError, SysErrorMessage(GetLastError)]);
      Exit;
    end;

    LSecDesc := AllocMem(LBytesNeeded);
    try
      // Second call to get the actual security descriptor
      if not GetFileSecurity(PChar(ToLongPath(ADirectory)), DACL_SECURITY_INFORMATION,
         LSecDesc, LBytesNeeded, LBytesNeeded) then
      begin
        AErrorDescription := Format('GetFileSecurity failed [%d]: %s',
          [GetLastError, SysErrorMessage(GetLastError)]);
        Exit;
      end;

      if not GetSecurityDescriptorDacl(LSecDesc, LDACLPresent, LDACL, LDefaulted) then
      begin
        AErrorDescription := Format('GetSecurityDescriptorDacl failed [%d]: %s',
          [GetLastError, SysErrorMessage(GetLastError)]);
        Exit;
      end;

      // No DACL present means full access to everyone — not a deny situation
      if not LDACLPresent or (LDACL = nil) then
        Exit;

      // Get number of ACEs in the DACL
      if not GetAclInformation(LDACL^, @LAclSizeInfo, SizeOf(LAclSizeInfo), TAclInformationClass(AclSizeInformation)) then
      begin
        AErrorDescription := Format('GetAclInformation failed [%d]: %s',
          [GetLastError, SysErrorMessage(GetLastError)]);
        Exit;
      end;

      for LAceIndex := 0 to LAclSizeInfo.AceCount - 1 do
      begin
        if not GetAce(LDACL^, LAceIndex, Pointer(LAceHeader)) then
          Continue;

        if LAceHeader^.AceType = ACCESS_DENIED_ACE_TYPE then
        begin
          LAccessDeniedAce := PACCESS_DENIED_ACE(LAceHeader);

          // Get the SID name for reporting who is denied
          var LSIDName: array[0..255] of Char;
          var LDomainName: array[0..255] of Char;
          var LSIDNameLen: DWORD := SizeOf(LSIDName);
          var LDomainNameLen: DWORD := SizeOf(LDomainName);
          var LSIDNameUse: SID_NAME_USE;

          if LookupAccountSid(nil, @LAccessDeniedAce^.SidStart,
             LSIDName, LSIDNameLen, LDomainName, LDomainNameLen, LSIDNameUse) then
            AErrorDescription := AErrorDescription +
              Format('DENY ACE found for: %s\%s  ', [LDomainName, LSIDName])
          else
            AErrorDescription := AErrorDescription + 'DENY ACE found for unknown SID  ';

          Result := True;
        end;
      end;

    finally
      FreeMem(LSecDesc);
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('HasExplicitDenyACE', ADirectory, E, AErrorDescription);
  end;
end;

procedure TFileRightsChecker.GetFilesAndDirs(const ADirectory: string; const AFiles, ADirectories: TStringList; const AClearLists: Boolean = True);
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
          GetFilesAndDirs(LPath + LSearchRec.Name, AFiles, ADirectories, False);  // recurse
        end
        else
          AFiles.Add(LPath + LSearchRec.Name);

      until FindNext(LSearchRec) <> 0;
    finally
      FindClose(LSearchRec);
    end;
  except
    on E: Exception do
    begin
      // TODO: Log etc...
      raise;
    end;
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
    LFileHandle := CreateFile(
      PChar(ToLongPath(LTempFile)),
      GENERIC_WRITE,
      0,
      nil,
      CREATE_NEW,
      FILE_ATTRIBUTE_NORMAL,
      0);

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

      if not WriteFile(LFileHandle, LData[1], Length(LData), LBytesWritten, nil)
         or (LBytesWritten <> DWORD(Length(LData))) then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('WriteFile failed [%d]: %s', [LErrorCode, SysErrorMessage(LErrorCode)]);

        Exit;
      end;
    finally
      CloseHandle(LFileHandle);
    end;

    // --- Delete ---
    if not DeleteFile(LTempFile) then
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
    LDirHandle := CreateFile(
      PChar(ToLongPath(ADirectory)),
      GENERIC_READ,
      FILE_SHARE_READ or FILE_SHARE_WRITE,
      nil,
      OPEN_EXISTING,
      FILE_FLAG_BACKUP_SEMANTICS,  // required to open a directory handle
      0);

    if LDirHandle = INVALID_HANDLE_VALUE then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('Directory read rights check failed [%d]: %s',
        [LErrorCode, SysErrorMessage(LErrorCode)]);
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

    LFileHandle := CreateFile(
      PChar(ToLongPath(AFileName)),
      GENERIC_EXECUTE,
      FILE_SHARE_READ,
      nil,
      OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL,
      0);
    try
      if LFileHandle = INVALID_HANDLE_VALUE then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('Execute rights check failed [%d]: %s',
          [LErrorCode, SysErrorMessage(LErrorCode)]);

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

    LFileHandle := CreateFile(
      PChar(ToLongPath(AFileName)),
      LAccessMode,
      FILE_SHARE_READ,
      nil,
      OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL,
      0);

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

function TFileRightsChecker.ToLongPath(const ADirectory: string): string;
const
  LONG_PATH_PREFIX     = '\\?\';
  UNC_PREFIX           = '\\';
  LONG_UNC_PATH_PREFIX = '\\?\UNC\';
begin
  if not FOpenFilesLongFileAndPathNameSupport then
    Exit(ADirectory);

  // Already prefixed
  if ADirectory.StartsWith(LONG_PATH_PREFIX) then
    Exit(ADirectory);

  // UNC path e.g. \\server\share
  if ADirectory.StartsWith(UNC_PREFIX) then
    Result := LONG_UNC_PATH_PREFIX + ADirectory.Substring(2)
  else
    Result := LONG_PATH_PREFIX + ADirectory;
end;

procedure TFileRightsChecker.Execute(const ADirectory: string; const ACheckWriteRights: Boolean);
begin
  var LFiles := TStringList.Create;
  var LSubDirectories := TStringList.Create;

  try
    InitializeDirectoriesAndFiles(ADirectory, LFiles, LSubDirectories);

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
    if IsRunningElevated then
      Exit;

    if IsProcess64Bit then
      Exit;

    LSystemDrive := GetEnvironmentVariable('SystemDrive');

    // Build the corresponding VirtualStore path
    if not ADirectory.StartsWith(LSystemDrive, True) then
      Exit;

    LVirtualStorePath := GetEnvironmentVariable('LOCALAPPDATA') +
      '\VirtualStore' + ADirectory.Substring(Length(LSystemDrive));

    // Only report if VirtualStore path actually exists and has content
    if DirectoryExists(LVirtualStorePath) and not IsDirectoryEmpty(LVirtualStorePath) then
    begin
      AErrorDescription := Format('Active UAC virtualization detected — files may be ' +
        'redirected to: %s', [LVirtualStorePath]);
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
      AErrorDescription := Format('OpenProcessToken failed [%d]: %s',
        [LErrorCode, SysErrorMessage(LErrorCode)]);
      Exit;
    end;
    try
      if not LookupPrivilegeValue(nil, PChar(APrivilegeName), LLUID) then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('LookupPrivilegeValue failed for "%s" [%d]: %s',
          [APrivilegeName, LErrorCode, SysErrorMessage(LErrorCode)]);
        Exit;
      end;

      LPrivilegeSet.PrivilegeCount := 1;
      LPrivilegeSet.Control := PRIVILEGE_SET_ALL_NECESSARY;
      LPrivilegeSet.Privilege[0].Luid := LLUID;
      LPrivilegeSet.Privilege[0].Attributes := 0;

      if not PrivilegeCheck(LTokenHandle, LPrivilegeSet, LHasPrivilege) then
      begin
        LErrorCode := GetLastError;
        AErrorDescription := Format('PrivilegeCheck failed [%d]: %s',
          [LErrorCode, SysErrorMessage(LErrorCode)]);
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
      AErrorDescription := Format('GetFileAttributes failed [%d]: %s',
        [LErrorCode, SysErrorMessage(LErrorCode)]);
      Exit;
    end;

    if (LAttributes and FILE_ATTRIBUTE_REPARSE_POINT) <> 0 then
    begin
      AErrorDescription := Format('Path is a reparse point (junction or symlink) — ' +
        'target location may have different access rights: %s', [ADirectory]);
      Result := True;
    end;
  except
    on E: Exception do
      GetExceptionErrorDescription('IsReparsePoint', ADirectory, E, AErrorDescription);
  end;
end;

procedure TFileRightsChecker.GetExceptionErrorDescription(const AErrorMethod, AFileSystemItem: string; const AException: Exception; var AErrorDescription: string);
begin
  AErrorDescription := 'Exception ' + AException.ClassName + ' occurred at ' + AErrorMethod.QuotedString('"')
    + ' with message: ' + AException.Message.QuotedString('"') + '. While checkinf file system item: '
    + AFileSystemItem.QuotedString('"');
end;

function TFileRightsChecker.GetFileIntegrityLevel(const ADirectory: string; var AErrorDescription: string): string;
const
  LABEL_SECURITY_INFORMATION = $10;
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
    LFileHandle := CreateFile(PChar(ToLongPath(ADirectory)),
      READ_CONTROL, FILE_SHARE_READ or FILE_SHARE_WRITE, nil, OPEN_EXISTING,
      FILE_FLAG_BACKUP_SEMANTICS, 0);

    if LFileHandle = INVALID_HANDLE_VALUE then
    begin
      LErrorCode := GetLastError;
      AErrorDescription := Format('CreateFile failed for integrity level check [%d]: %s',
        [LErrorCode, SysErrorMessage(LErrorCode)]);
      Exit;
    end;
    try
      // First call to get buffer size
      GetKernelObjectSecurity(LFileHandle, LABEL_SECURITY_INFORMATION,
        nil, 0, LBytesNeeded);

      if LBytesNeeded = 0 then
      begin
        Result := 'No integrity label — defaults to Medium';
        Exit;
      end;

      LSecDesc := AllocMem(LBytesNeeded);
      try
        if not GetKernelObjectSecurity(LFileHandle, LABEL_SECURITY_INFORMATION,
           LSecDesc, LBytesNeeded, LBytesNeeded) then
        begin
          LErrorCode := GetLastError;
          AErrorDescription := Format('GetKernelObjectSecurity failed [%d]: %s',
            [LErrorCode, SysErrorMessage(LErrorCode)]);
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
    if IsRunningElevated then
      AErrorDescription := AErrorDescription +
        ' [Process IS elevated — likely explicit DENY ACE or EFS encryption]'
    else
      AErrorDescription := AErrorDescription +
        ' [Process is NOT elevated — re-run as Administrator to confirm. If that succeeds, likely cause is UAC token filtering,'
        + ' explicit DENY ACE, or EFS encryption]'
  end;
end;

constructor TFileRightsChecker.Create(const AOpenFilesLongFileAndPathNameSupport: Boolean = True;
      const ACheckProcessBackupPrivileges: Boolean = False);
begin
  inherited Create;

  FReadOnlyStatistics := TStatistics.Create;
  FReadWriteStatistics := TStatistics.Create;
  FErrors := TErrorItemCollection.Create;
  FOpenFilesLongFileAndPathNameSupport := AOpenFilesLongFileAndPathNameSupport;
  FCheckProcessBackupPrivileges := ACheckProcessBackupPrivileges;
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
    var LErrorDescription: string := '';

    if IsDirectoryUnderUACVirtualization(LSubDirectory, LErrorDescription) then
      LogError(LSubDirectory, frcUACVirtualization, LErrorDescription)
    else if ACheckWriteRights then
    begin
      if not TestDirectoryWriteRights(LSubDirectory, LErrorDescription) then
        LogError(LSubDirectory, frcDirectoryNotWritable, LErrorDescription)
      else
        FReadWriteStatistics.AddCheckedDirectory;
    end
    else
    begin
      if not TestDirectoryReadRights(LSubDirectory, LErrorDescription) then
        LogError(LSubDirectory, frcDirectoryNotReadable , LErrorDescription)
      else
        FReadOnlyStatistics.AddCheckedDirectory;
    end;
  end;
end;

procedure TFileRightsChecker.DoFileChecks(const AFiles: TStringList; const ACheckWriteRights: Boolean);
begin
  for var LCurrentFile in AFiles do
  begin
    var LErrorDescription: string := '';

    if IsReparsePoint(LCurrentFile, LErrorDescription) then
      LogError(LCurrentFile, frcIsReparsePoint, LErrorDescription)
    else if not TestOpenFileRights(LCurrentFile, ACheckWriteRights, LErrorDescription) then
    begin
      var LDenyDescription: string := '';

      if HasExplicitDenyACE(LCurrentFile, LDenyDescription) then
        LErrorDescription := LErrorDescription + ' | ' + LDenyDescription
      else if HasEmptyDACL(LCurrentFile, LDenyDescription) then
        LErrorDescription := LErrorDescription + ' | ' + LDenyDescription;

      var LIntegrityLevelErrorDescription: string;
      var LIntegrityLevel: string := GetFileIntegrityLevel(LCurrentFile, LIntegrityLevelErrorDescription);
      if LIntegrityLevel <> '' then
        LErrorDescription := LErrorDescription + Format(' | Integrity level: %s', [LIntegrityLevel]) +
        IfThen(LIntegrityLevelErrorDescription.IsEmpty, '', ' - ' + LIntegrityLevelErrorDescription);

      LogError(LCurrentFile, TFileRightErrorType(IfThen(ACheckWriteRights, Integer(frcFileNotWritable), Integer(frcFileNotReadable))), LErrorDescription);

      var LExtension := ExtractFileExt(LCurrentFile);

      if MatchText(LExtension, ['.exe', '.dll']) then
        if not TestExecuteRights(LCurrentFile, LErrorDescription) then
          LogError(LCurrentFile, frcUserHasNoExecuteRightsForFile, LErrorDescription);
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
    frcMissingPrivilege: Result := 'Process is missing required privilege';
    frcFileNotReadable: Result := 'File not readable';
    frcFileNotWritable: Result := 'File not writable';
    frcUserHasNoExecuteRightsForFile: Result := 'User has no execute rights';
    frcIsReparsePoint: Result := 'Path is a reparse point or symlink';
    frcUACVirtualization: Result := 'Path is under UAC virtualization';
    frcDirectoryNotReadable: Result := 'Directory not readable';
    frcDirectoryNotWritable: Result := 'Directory not writable';
    else
      Result := Format('Unknown error type (%d)', [Ord(FErrorType)]);
  end;
end;

{ TErrorItemCollection }

procedure TErrorItemCollection.Add(const AItem: TErrorItem);
begin
  FErrorItems.Add(AItem)
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
