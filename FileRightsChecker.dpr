program FileRightsChecker;

uses
  Vcl.Forms,
  FRCForm.Main in 'FRCForm.Main.pas' {FRCMainForm},
  FRCUnit.FileRightsChecker in 'FRCUnit.FileRightsChecker.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFRCMainForm, FRCMainForm);
  Application.Run;
end.
