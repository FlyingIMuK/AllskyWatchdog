program AllskyWatchdog;

uses
  Forms,
  fAllskyWatchdog in 'fAllskyWatchdog.pas' {frmAllskyWatchdog},
  fMsgWin in '..\fMsgWin.pas' {frmMessageWindow},
  mGlobProc_B in '..\AAA_SharedFiles\mGlobProc_B.pas',
  mMUKSunCalc_B in '..\AAA_SharedFiles\mMUKSunCalc_B.pas',
  mApiFunctions_B in '..\AAA_SharedFiles\mApiFunctions_B.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmAllskyWatchdog, frmAllskyWatchdog);
  Application.CreateForm(TfrmMessageWindow, frmMessageWindow);
  Application.Run;
end.
