program FileRightsChecker;

uses
  Vcl.Forms,
  FRCForm.Main in 'Source\Forms\FRCForm.Main.pas' {FRCMainForm},
  FRCUnit.FileRightsChecker in 'Source\Units\FRCUnit.FileRightsChecker.pas',
  FRCUnit.Settings in 'Source\Units\FRCUnit.Settings.pas',
  FRCUnit.WinAPI in 'Source\Units\FRCUnit.WinAPI.pas',
  FRCUnit.ProgressThrottle in 'Source\Units\FRCUnit.ProgressThrottle.pas',
  FRCUnit.Statistics in 'Source\Units\FRCUnit.Statistics.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFRCMainForm, FRCMainForm);
  Application.Run;
end.
