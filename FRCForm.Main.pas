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
    procedure LoadSettings;
    procedure SaveSettings;
    procedure Log(const ALogLine: string; const AIndent: Integer = 0);
  end;

var
  FRCMainForm: TFRCMainForm;

implementation

{$R *.dfm}

uses
  FRCUnit.FileRightsChecker;

const
  SETTINGS_FILENAME = 'FRCSettings.json';

procedure TFRCMainForm.FormCreate(Sender: TObject);
begin
  FSettings := TFRCSettings.Create;
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

  // Time to refresh the screen, most likely not needed.
  Application.ProcessMessages;

  var LFileAccessCheck := TFileRightsChecker.Create(CheckBoxOpenFilesLongFileAndPathNameSupport.Checked,
    CheckBoxProcessBackupPrivileges.Checked, CheckBoxRunDirectoryGetEffectiveRightsShortfallTests.Checked,
    CheckBoxRunFileGetEffectiveRightsShortfallTests.Checked, CheckBoxRunCurrentUserIsOwnerTests.Checked);
  try
    LFileAccessCheck.Execute(EditReadWriteChecks.Text, True);
    LFileAccessCheck.Execute(EditReadOnlyCheck.Text, False);

    if LFileAccessCheck.Errors.Count = 0 then
      Log('No errors found')
    else
    begin
      for var LIndex := 0 to LFileAccessCheck.Errors.Count - 1 do
      begin
        var LError := LFileAccessCheck.Errors[LIndex];

        Log(LError.ErrorTypeStr + '- '  + LError.FileSystemItem.QuotedString('"') + ' - With error: ' + LError.ErrorDescription, 1);
      end;
    end;

    Log('Statistics:');
    Log('- ' + LFileAccessCheck.ReadOnlyStatistics.DirectoriesChecked.ToString + ' directories are readable (as should)', 1);
    Log('- ' + LFileAccessCheck.ReadOnlyStatistics.FilesChecked.ToString + ' files could be opened in read only-mode', 1);
    Log('- ' + LFileAccessCheck.ReadWriteStatistics.DirectoriesChecked.ToString + ' directories are writable (as should)', 1);
    Log('- ' + LFileAccessCheck.ReadWriteStatistics.FilesChecked.ToString + ' files could be opened in read write-mode', 1);
    Log('', 1);
  finally
    LFileAccessCheck.Free;
    LAction.Enabled := True;
    Screen.Cursor := crDefault;
  end;
end;

procedure TFRCMainForm.FormShow(Sender: TObject);
begin
  LoadSettings;
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
