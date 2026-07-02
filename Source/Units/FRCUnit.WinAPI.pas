unit FRCUnit.WinAPI;

// Win32 / advapi32 declarations that either aren't exposed by Winapi.Windows /
// Winapi.AccCtrl in all Delphi versions, or that are project-specific aliases.
// Keeping them in one place keeps the rest of the project free of raw imports.

interface

uses
  Winapi.Windows, Winapi.AccCtrl;

const
  // ACE type values used in the AceType field of ACE headers.
  ACCESS_ALLOWED_ACE_TYPE          = Byte($0);
  ACCESS_DENIED_ACE_TYPE           = Byte($1);
  SYSTEM_AUDIT_ACE_TYPE            = Byte($2);
  SYSTEM_ALARM_ACE_TYPE            = Byte($3);
  ACCESS_ALLOWED_COMPOUND_ACE_TYPE = Byte($4);
  ACCESS_ALLOWED_OBJECT_ACE_TYPE   = Byte($5);
  ACCESS_DENIED_OBJECT_ACE_TYPE    = Byte($6);
  SYSTEM_AUDIT_OBJECT_ACE_TYPE     = Byte($7);

  // ACE header flag: the ACE applies only to inherited children, not the object itself.
  INHERIT_ONLY_ACE_FLAG            = Byte($08);

  // SECURITY_INFORMATION bit for the mandatory integrity label (SACL flavour).
  LABEL_SECURITY_INFORMATION       = $10;

  // Composite file-access masks from <winnt.h>. Not always exposed by Winapi.Windows,
  // so spelled out here. STANDARD_RIGHTS_READ/WRITE = $00020000, SYNCHRONIZE = $00100000.
  //   FILE_GENERIC_READ  = STANDARD_RIGHTS_READ  | FILE_READ_DATA  ($1) | FILE_READ_ATTRIBUTES  ($80)  | FILE_READ_EA  ($8)  | SYNCHRONIZE
  //   FILE_GENERIC_WRITE = STANDARD_RIGHTS_WRITE | FILE_WRITE_DATA ($2) | FILE_WRITE_ATTRIBUTES ($100) | FILE_WRITE_EA ($10) | FILE_APPEND_DATA ($4) | SYNCHRONIZE
  FILE_GENERIC_READ                = $00120089;
  FILE_GENERIC_WRITE               = $00120116;

  // Attribute bits for offline / cloud-placeholder files (OneDrive, Dropbox, HSM).
  // The RECALL_* bits are Windows 10 era and missing from older Winapi.Windows.
  FILE_ATTRIBUTE_OFFLINE_BIT             = $00001000;
  FILE_ATTRIBUTE_RECALL_ON_OPEN          = $00040000;
  FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS   = $00400000;

  // GetVolumeInformation file-system flag: every write on the volume fails.
  FILE_READ_ONLY_VOLUME                  = $00080000;

  // TOKEN_INFORMATION_CLASS value for the process integrity level
  // (TokenIntegrityLevel) — enum member missing from some older RTLs, so we use
  // the raw ordinal with a hard cast.
  TOKEN_INTEGRITY_LEVEL_INFO_CLASS       = 25;

  // WinError codes referenced in failure-hint texts; spelled out in case the RTL
  // version in use predates them.
  ERROR_DISK_QUOTA_EXCEEDED              = 1295;
  ERROR_CANT_ACCESS_FILE                 = 1920;

  // Standard access right DELETE from <winnt.h>. Winapi.Windows declares it as
  // DELETE, but that identifier collides with the intrinsic System.Delete procedure
  // in Delphi, so the project uses this alias.
  DELETE_ACCESS_RIGHT                    = $00010000;

type
  TACE_HEADER = record
    AceType:  Byte;
    AceFlags: Byte;
    AceSize:  Word;
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

  // Not exposed in Winapi.Windows in all Delphi versions — declare it here so we
  // don't depend on the RTL having it.
  function CheckTokenMembership(TokenHandle: THandle; SidToCheck: PSID; var IsMember: BOOL): BOOL; stdcall;
    external 'advapi32.dll' name 'CheckTokenMembership';

  // Same story for GetEffectiveRightsFromAclW — Winapi.AccCtrl declares the
  // TRUSTEE_W types but doesn't always declare the function import.
  function GetEffectiveRightsFromAcl(const pacl: TACL; const pTrustee: TRUSTEE_W; var AccessRights: ACCESS_MASK): DWORD; stdcall;
    external 'advapi32.dll' name 'GetEffectiveRightsFromAclW';

implementation

end.
