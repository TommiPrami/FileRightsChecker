unit FRCForm.Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, System.Actions, Vcl.ActnList, Vcl.ExtCtrls;

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
    procedure ActionRunExecute(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FRCMainForm: TFRCMainForm;

implementation

{$R *.dfm}

uses
  FRCUnit.FileRightsChecker;

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

end.
