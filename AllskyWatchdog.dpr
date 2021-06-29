program AllskyWatchdog;

uses
  Forms,
  fAllskyWatchdog in 'fAllskyWatchdog.pas' {frmDataWatchdog},
  fMsgWin in '..\fMsgWin.pas' {frmMessageWindow},
  mGlobProc_B in '..\AAA_SharedFiles\mGlobProc_B.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmDataWatchdog, frmDataWatchdog);
  Application.CreateForm(TfrmMessageWindow, frmMessageWindow);
  Application.Run;
end.
