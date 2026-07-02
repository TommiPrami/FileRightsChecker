unit FRCForm.Main;

interface

uses
  Winapi.Messages, Winapi.Windows, System.Actions, System.Classes, System.SysUtils, System.Variants, Vcl.ActnList,
  Vcl.Controls, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Forms, Vcl.Graphics, Vcl.StdCtrls, FRCUnit.Settings;

type
  TFRCMainForm = class(TForm)
    ActionList: TActionList;
    ActionRun: TAction;
    ButtonRun: TButton;
    CheckBoxOpenFilesLongFileAndPathNameSupport: TCheckBox;
    EditReadOnlyCheck: TEdit;
    EditReadWriteChecks: TEdit;
    LabelMustHaveReadRights: TLabel;
    LabelMustHaveWriteRights: TLabel;
    MemoLog: TMemo;
    PanelButtons: TPanel;
    PanelLeft: TPanel;
    PanelLog: TPanel;
    PanelTop: TPanel;
    CheckBoxProcessBackupPrivileges: TCheckBox;
    CheckBoxRunFileGetEffectiveRightsShortfallTests: TCheckBox;
    CheckBoxRunDirectoryGetEffectiveRightsShortfallTests: TCheckBox;
    CheckBoxRunCurrentUserIsOwnerTests: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ActionRunExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
  strict private
    FSettings: TFRCSettings;
    function GetExeFileVersion(const ASegments: Integer): string;
    procedure LoadSettings;
    procedure SaveSettings;
    procedure Log(const ALogLine: string; const AIndent: Integer = 0);
    procedure ShowProgress(const AProgressText: string);
  end;

var
  FRCMainForm: TFRCMainForm;

implementation

{$R *.dfm}

uses
  FRCUnit.FileRightsChecker, FRCUnit.ProgressThrottle, FRCUnit.Statistics;

const
  SETTINGS_FILENAME = 'FRCSettings.json';

procedure TFRCMainForm.FormCreate(Sender: TObject);
begin
  FSettings := TFRCSettings.Create;

  Caption := Caption + ' v' + GetExeFileVersion(3);
end;

procedure TFRCMainForm.FormDestroy(Sender: TObject);
begin
  SaveSettings;

  FSettings.Free;
end;

procedure TFRCMainForm.ActionRunExecute(Sender: TObject);
begin
  var LAction := Sender as TAction;

  LAction.Enabled := False;
  Screen.Cursor := crHourGlass;
  MemoLog.Clear;

  // Time to refresh the screen, most likely not needed.
  Application.ProcessMessages;

  var LFileAccessCheck := TFileRightsChecker.Create(CheckBoxOpenFilesLongFileAndPathNameSupport.Checked,
    CheckBoxProcessBackupPrivileges.Checked, CheckBoxRunDirectoryGetEffectiveRightsShortfallTests.Checked,
    CheckBoxRunFileGetEffectiveRightsShortfallTests.Checked, CheckBoxRunCurrentUserIsOwnerTests.Checked);
  // Hand the throttle the current caption — it will prepend it to every progress
  // update and restore the bare caption once 100% is reached.
  var LThrottle := TProgressThrottle.Create(Caption);
  try
    LThrottle.OnShowProgress := ShowProgress;
    LFileAccessCheck.OnTest := LThrottle.HandleTest;

    // Who ran the scan and how — without this, support cannot interpret the results.
    Log(LFileAccessCheck.RunContextDescription);
    Log('');

    // Queue both passes first so RunPreparedPasses can see the full picture and
    // emit one continuous 0..100% sweep with cumulative test and error counts.
    LFileAccessCheck.PreparePass(EditReadWriteChecks.Text, True);
    LFileAccessCheck.PreparePass(EditReadOnlyCheck.Text, False);
    LFileAccessCheck.RunPreparedPasses;

    if LFileAccessCheck.Errors.Count = 0 then
      Log('No errors found')
    else
    begin
      for var LIndex := 0 to LFileAccessCheck.Errors.Count - 1 do
      begin
        var LError := LFileAccessCheck.Errors[LIndex];

        Log('[' + LError.SeverityStr + '] ' + LError.ErrorTypeStr + ' - ' + LError.FileSystemItem.QuotedString('"')
          + ' - With error: ' + LError.ErrorDescription, 1);
      end;

      Log('');
      Log(Format('Findings: %d errors, %d warnings, %d info', [LFileAccessCheck.Errors.ErrorCount,
        LFileAccessCheck.Errors.CountBySeverity(esWarning), LFileAccessCheck.Errors.CountBySeverity(esInfo)]));
    end;

    Log('Statistics:');
    Log('- ' + LFileAccessCheck.ReadOnlyStatistics.DirectoriesChecked.ToString + ' directories are readable (as should)', 1);
    Log('- ' + LFileAccessCheck.ReadOnlyStatistics.FilesChecked.ToString + ' files could be opened in read only-mode', 1);
    Log('- ' + LFileAccessCheck.ReadWriteStatistics.DirectoriesChecked.ToString + ' directories are writable (as should)', 1);
    Log('- ' + LFileAccessCheck.ReadWriteStatistics.FilesChecked.ToString + ' files could be opened in read write-mode', 1);
    Log('- ' + LFileAccessCheck.ReadWriteStatistics.FilesOpenedExclusively.ToString + ' files could be opened in exclusive mode', 1);
    Log('', 1);
  finally
    // Restore the title even if the run died with an exception — otherwise the
    // caption stays stuck at the last progress text.
    Caption := LThrottle.CaptionPrefix;
    LThrottle.Free;
    LFileAccessCheck.Free;
    LAction.Enabled := True;
    Screen.Cursor := crDefault;
  end;
end;

procedure TFRCMainForm.ShowProgress(const AProgressText: string);
begin
  // Throttler decides cadence and supplies the full text (caption prefix already
  // included). We just paint it on the title bar and pump messages so the form
  // stays responsive during long scans on the UI thread.
  Caption := AProgressText;
  Application.ProcessMessages;
end;

procedure TFRCMainForm.FormShow(Sender: TObject);
begin
  LoadSettings;
end;

function TFRCMainForm.GetExeFileVersion(const ASegments: Integer): string;
var
  LExeName: string;
  LDummyHandle: DWORD;
  LInfoSize: DWORD;
  LBuffer: TBytes;
  LFixed: PVSFixedFileInfo;
  LFixedLen: UINT;
  LCount: Integer;
begin
  Result := '';

  // Clamp the segment count to the documented 1..4 range so callers can't
  // accidentally print "1.2.2.6.0..." or an empty string.
  LCount := ASegments;
  if LCount < 1 then
    LCount := 1
  else if LCount > 4 then
    LCount := 4;

  LExeName := ParamStr(0);
  LInfoSize := GetFileVersionInfoSize(PChar(LExeName), LDummyHandle);
  if LInfoSize = 0 then
    Exit;

  SetLength(LBuffer, LInfoSize);
  if not GetFileVersionInfo(PChar(LExeName), 0, LInfoSize, LBuffer) then
    Exit;

  // '\' returns the root VS_FIXEDFILEINFO block — language-independent, which is
  // what we want for the numeric version (the localized "FileVersion" string is
  // a separate VarFileInfo subblock).
  if not VerQueryValue(LBuffer, '\', Pointer(LFixed), LFixedLen) then
    Exit;

  // VS_FIXEDFILEINFO packs the four 16-bit components into two DWORDs:
  //   dwFileVersionMS = (major << 16) | minor
  //   dwFileVersionLS = (release << 16) | build
  Result := IntToStr(HiWord(LFixed^.dwFileVersionMS));

  if LCount >= 2 then
    Result := Result + '.' + IntToStr(LoWord(LFixed^.dwFileVersionMS));

  if LCount >= 3 then
    Result := Result + '.' + IntToStr(HiWord(LFixed^.dwFileVersionLS));

  if LCount >= 4 then
    Result := Result + '.' + IntToStr(LoWord(LFixed^.dwFileVersionLS));
end;

procedure TFRCMainForm.LoadSettings;
begin
  FSettings.LoadFromFile(SETTINGS_FILENAME);

  EditReadOnlyCheck.Text := FSettings.ReadOnlyDirectoriesAsString;
  EditReadWriteChecks.Text := FSettings.ReadWriteDirectoriesAsString;
end;

procedure TFRCMainForm.Log(const ALogLine: string; const AIndent: Integer = 0);
begin
  MemoLog.Lines.Add(StringOfChar(' ', AIndent * 2) + ALogLine);
end;

procedure TFRCMainForm.SaveSettings;
begin
  FSettings.ParseReadOnlyDirectoriesFromString(EditReadOnlyCheck.Text);
  FSettings.ParseReadWriteDirectoriesFromString(EditReadWriteChecks.Text);

  FSettings.SaveToFile(SETTINGS_FILENAME);
end;

end.
