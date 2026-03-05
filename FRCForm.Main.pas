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
    CheckBoxCheckProcessBackupPrivileges: TCheckBox;
    CheckBoxOpenFilesLongFileAndPathNameSupport: TCheckBox;
    EditReadOnlyCheck: TEdit;
    EditReadWrtiteChecks: TEdit;
    LabelMustHaveReadRights: TLabel;
    LabelMustHaveWriteRights: TLabel;
    MemoLog: TMemo;
    PanelButtons: TPanel;
    PanelLeft: TPanel;
    PanelLog: TPanel;
    PanelTop: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ActionRunExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
  strict private
    FSettings: TFRCSettings;
    procedure LoadSettings;
    procedure SaveSettings;
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
    CheckBoxCheckProcessBackupPrivileges.Checked);
  try
    LFileAccessCheck.Execute(EditReadWrtiteChecks.Text, True);
    LFileAccessCheck.Execute(EditReadOnlyCheck.Text, False);

    if LFileAccessCheck.Errors.Count = 0 then
      MemoLog.Lines.Add('No errors found')
    else
    begin
      for var LIndex := 0 to LFileAccessCheck.Errors.Count - 1 do
      begin
        var LError := LFileAccessCheck.Errors[LIndex];

        MemoLog.Lines.Add(LError.ErrorTypeStr + ' - '  + LError.FileSystemItem.QuotedString('"') + ' - With error: ' + LError.ErrorDescription)
      end;
    end;

    MemoLog.Lines.Add('Statistics:');
    MemoLog.Lines.Add('  - ' + LFileAccessCheck.ReadOnlyStatistics.FilesChecked.ToString + ' files could be opened in read only-mode');
    MemoLog.Lines.Add('  - ' + LFileAccessCheck.ReadWriteStatistics.DirectoriesChecked.ToString + ' directories are writable (as should)');
    MemoLog.Lines.Add('  - ' + LFileAccessCheck.ReadWriteStatistics.FilesChecked.ToString + ' files could be opened in read write-mode');
    MemoLog.Lines.Add('  ');
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

  EditReadOnlyCheck.Text := FSettings.ReadOnlyDirectoriesStr;
  EditReadWrtiteChecks.Text := FSettings.ReadWriteDirectoriesStr;
end;

procedure TFRCMainForm.SaveSettings;
begin
  FSettings.Clear;

  FSettings.AddReadOnlyDirectories(EditReadOnlyCheck.Text);
  FSettings.AddReadWriteDirectories(EditReadWrtiteChecks.Text);

  FSettings.SaveToFile(SETTINGS_FILENAME);
end;

end.
