program FileRightsChecker;

uses
  Vcl.Forms,
  FRCForm.Main in 'FRCForm.Main.pas' {FRCMainForm},
  FRCUnit.FileRightsChecker in 'FRCUnit.FileRightsChecker.pas',
  FRCUnit.Settings in 'FRCUnit.Settings.pas',
  FRCUnit.WinAPI in 'FRCUnit.WinAPI.pas',
  FRCUnit.ProgressThrottle in 'FRCUnit.ProgressThrottle.pas',
  FRCUnit.Statistics in 'FRCUnit.Statistics.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFRCMainForm, FRCMainForm);
  Application.Run;
end.
