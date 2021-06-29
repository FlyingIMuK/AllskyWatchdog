unit fAllskyWatchdog;
(* Ausgangsversion von Ulli Finke
  Checkt alle 5 Minuten die
  Status-Datei
  'lastLAST_AUTO_TESTS.txt' und
  schreibt die letzten 3 Zeilen in eine
  Datei ins Web: meteo/safstat.txt
  Wenn eine Station zu alt ist, wird eine
  email geschickt.


  Überarbeitet von Holger Schilke 12.11.2008
  - Die Komponente TIdSMTP wird verwendet, statt TSmtpCli aus externem Package
  - Die eMail-Liste läßt sich editieren.
*)

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, shellapi, Vcl.Mask, JvExMask, JvToolEdit, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdMessageClient, IdSMTPBase, IdSMTP, Vcl.Grids, ColorGrid, Vcl.ComCtrls,
  HSMultiAlarmClock, Vcl.ExtCtrls, Vcl.Samples.Spin, Vcl.Buttons,
  Inifiles, math,system.DateUtils;

const
  TimeOut = 5; // Minute (Max. Dateialter)

type
  TfrmDataWatchdog = class(TForm)
    bbnClose: TBitBtn;
    speInterval: TSpinEdit;
    Panel1: TPanel;
    chkEmail: TCheckBox;
    Mail: TButton;
    lbEmail: TListBox;
    bbnAdd: TBitBtn;
    bbnDelete: TBitBtn;
    IdSMTP1: TIdSMTP;
    HSMultiAlarmClock1: THSMultiAlarmClock;
    StatusBar1: TStatusBar;
    Label2: TLabel;
    chkAutostart: TCheckBox;
    cgdFileList: TXColorGrid;
    bbnAddPath: TBitBtn;
    bbnCheck: TBitBtn;
    Label1: TLabel;
    slePrefix: TEdit;
    Label3: TLabel;
    sleExt: TEdit;
    Label6: TLabel;
    speTimeout: TSpinEdit;
    deAllskyPath: TJvDirectoryEdit;
    StringGrid1: TStringGrid;
    (* procedure SmtpCli1RequestDone(Sender: TObject; RqType: TSmtpRequest;
      ErrorCode: Word); *)
    procedure FormCreate(Sender: TObject);
    procedure MailClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bbnHelpClick(Sender: TObject);
    procedure bbnCloseClick(Sender: TObject);
    procedure bbnAddClick(Sender: TObject);
    procedure bbnDeleteClick(Sender: TObject);
    procedure CheckData;

    procedure NextAlarm;
    procedure chkAutostartClick(Sender: TObject);
    procedure HSMultiAlarmClock1Alarm(Sender: TObject);
    procedure FormShow(Sender: TObject);
    function dCheckLastDate(sFileName: string): TDateTime;
    procedure FailedMessage(sMessage: string);
    procedure SendMailToList(sMessage: string);
    procedure bbnAddPathClick(Sender: TObject);
    procedure feDataFileAfterDialog(Sender: TObject; var Name: string;
      var Action: Boolean);
    procedure bbnCheckClick(Sender: TObject);
    procedure btnTestPowermanagerClick(Sender: TObject);
    procedure ReconnectLogger;
    function bIsAllskyDir(sFN: string): Boolean;
function dAllskyFNToDate(sAllskyFN: string): TDateTime;
  private
    _bMailSent: Boolean;
    hMainHandle, hErrorWindow: HWND;
    _sPath, _sFN, _sExt: string;

    { Private-Deklarationen }
  public
    { Public-Deklarationen }
    aLastMessage: packed array [1 .. 3] of TDateTime;
    bDoWarn, bAltWarnServer, bAltWarnClients: Boolean;
    aWarned: packed array [1 .. 3] of Boolean;
    // Für diese Station ist schon eine Warnung raus gegangen
    sSpectromatIniFile, sGlobLogFile, sGlobAktDir, sGlobIniFile: string;
  end;

var
  f: text;
  Mutex: THandle;
  h: HWND;
  frmDataWatchdog: TfrmDataWatchdog;

implementation

uses fMsgWin, mGlobProc_B;

{$R *.DFM}
(* ****************************************************************** *)

type
  s3_array = array [1 .. 3] of string;

  (* ****************************************************************** *)

procedure FileOperation(const source, dest: string; op, flags: Integer);
var
  shf: TSHFileOpStruct;
  s1, s2: string;
begin
  FillChar(shf, SizeOf(shf), #0);
  s1 := source + #0#0;
  s2 := dest + #0#0;
  shf.Wnd := 0;
  shf.wFunc := op;
  shf.pFrom := PCHAR(s1);
  shf.pTo := PCHAR(s2);
  shf.fFlags := flags;
  SHFileOperation(shf);
end (* FileOperation *);

(* ****************************************************************** *)

procedure TfrmDataWatchdog.FormCreate(Sender: TObject);
var
  fIni: TIniFile;
  sDir, s, sExe, sName, sExtension: string;
  dDate: TDateTime;

  i: Integer;
  rect: TGridRect;
begin
  bDoWarn := true;
  bAltWarnClients := false;
  // datspinedit1.value := now;
  (* -------- *)
  FormatSettings.Decimalseparator := '.';
  FormatSettings.LongDateFormat := 'dd.mm.yyyy hh:nn:ss';
  FormatSettings.ShortTimeFormat := 'hh:nn';

  sExe := Application.ExeName;
  dDate := FileDateToDateTime(FileAge(sExe));
  Caption := 'Data-Watchdog  ' +
  // 'Version ' + sVersionsInfo(sExe) +
    ' (' + DateTimeToStr(dDate) + ')';

  sGlobAktDir := ExtractFilePath(sExe);
  AnalyseFileName(sExe, sGlobAktDir, sName, sExtension);
  sGlobIniFile := sGlobAktDir + '\' + sName + '.Ini';
  sGlobLogFile := sGlobAktDir + '\' + sName + '_Log.txt';
  (* -------- *)
  lWriteToErrorLog(sGlobLogFile, 'Application (Version ' + DateTimeToStr(dDate)
    + ') started');

  fIni := TIniFile.Create(sGlobIniFile);
  with fIni do
  begin
    speInterval.Value := ReadInteger('Global', 'Interval', 30);
    speTimeout.Value := ReadInteger('Global', 'Timeout', 30);

    chkAutostart.Checked := Readbool('Autostart', 'Enabled', false);
    // muss hinter speInterval stehen
    chkEmail.Checked := Readbool('Global', 'SendMail', false);
    i := 1;
    cgdFileList.ClearGrid;
    repeat

      s := ReadString('Files', 'Path' + inttostr(i), '');
      if s <> '' then
      begin
        with cgdFileList do
        begin
          cells[0, 0] := 'Nr';
          cells[1, 0] := 'Path';
          cells[2, 0] := 'File';
          cells[3, 0] := 'Ext';
          cells[4, 0] := 'LastTime';
          cells[5, 0] := 'Difference';
          cells[0, i] := inttostr(i);
          cells[1, i] := s;
          cells[2, i] := ReadString('Files', 'Prefix' + inttostr(i), '');
          cells[3, i] := ReadString('Files', 'Ext' + inttostr(i), '');
        end;
        cgdFileList.RowCount := cgdFileList.RowCount + 1;
      end;
      inc(i);
    until s = '';

    i := 0;
    repeat
      s := ReadString('Adresse', 'Nr' + inttostr(i), '');
      if s <> '' then
        lbEmail.Items.Add(s);
      inc(i);
    until s = '';
    // AnalyseFileName(feSpectromatExe.Text, sDir, sName, sExtension);
    (* sSpectromatIniFile := sDir + '\' + sName + '.Ini';
      deDataDir.Text := ReadString('Spectromat', 'DataDir', 'c:');
      rbDailySubDir.checked := ReadBool('Spectromat', 'DailySubdirectory', False);
      rbNSubDir.checked := ReadBool('Spectromat', 'SubdirNFiles', False);
      deDataDir.InitialDir := deDataDir.Text;
      fePowerManager.Text := ReadString('PowerManagement', 'ExeFile', 'C:');
      sleDevice.Text := ReadString('PowerManagement', 'Device', 'SkyScanner');
      sleSocket1.Text := ReadString('PowerManagement', 'Socket1', 'Socket1');
      //    chkSkyscannerUsed.Checked := ReadBool('SkcScanner', 'Used', False);
    *)
    Application.ProcessMessages;
    free;
  end;

  _bMailSent := false;

end;
(* ****************************************************************** *)

procedure TfrmDataWatchdog.MailClick(Sender: TObject);
begin
  SendMailToList('Data-Watchdog-Testmail');
end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.SendMailToList(sMessage: string);
var
  i: Integer;
begin
  for i := 0 to lbEmail.Items.Count - 1 do
    IdSMTP1.QuickSend('mailgate.uni-hannover.de', sMessage, lbEmail.Items[i],
      'schilke@muk.uni-hannover.de', sMessage);

end;
(* ****************************************************************** *)

procedure TfrmDataWatchdog.FormClose(Sender: TObject; var Action: TCloseAction);
var
  fIni: TIniFile;
  i: Integer;
begin
  fIni := TIniFile.Create(sGlobIniFile);
  with fIni do
  begin
    EraseSection('Adresse');
    for i := 0 to lbEmail.Items.Count - 1 do
      WriteString('Adresse', 'Nr' + inttostr(i), lbEmail.Items[i]);
    EraseSection('Files');
    for i := 1 to cgdFileList.RowCount - 1 do
      if cgdFileList.cells[1, i] <> '' then
      begin
        WriteString('Files', 'Path' + inttostr(i), cgdFileList.cells[1, i]);
        WriteString('Files', 'Prefix' + inttostr(i), cgdFileList.cells[2, i]);
        WriteString('Files', 'Ext' + inttostr(i), cgdFileList.cells[3, i]);
      end;

    WriteBool('Autostart', 'Enabled', chkAutostart.Checked);
    WriteInteger('Global', 'Interval', speInterval.Value);
    WriteInteger('Global', 'Timeout', speTimeout.Value);

    WriteBool('Global', 'SendMail', chkEmail.Checked);

    free;
  end;
  lWriteToErrorLog(sGlobLogFile, 'Application terminated');
end;
(* ****************************************************************** *)

procedure TfrmDataWatchdog.bbnHelpClick(Sender: TObject);
begin (*
    frmMessageWindow.sCaption := 'SafStat - Der Status von Safir';
    frmMessageWindow.sMessage := 'Safir-Status ' + scNewLine + scNewLine +
    'Checkt alle 5 Minuten die ' + scNewLine +
    'Status-Datei "lastLAST_AUTO_TESTS.txt" und' + scNewLine +
    'schreibt die letzten 3 Zeilen in eine' + scNewLine +
    'Datei ins Web: meteo/safstat.txt' + scNewLine +
    'Wenn eine Station zu alt ist, wird eine email geschickt.' + scNewLine +
    ' (UF, 30.03.2004)' + scNewLine +
    ' (HS, 12.11.2008)';

    frmMessageWindow.ShowModal; *)
end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.bbnCloseClick(Sender: TObject);
begin
  close;
end;
(* ****************************************************************** *)

procedure TfrmDataWatchdog.bbnAddClick(Sender: TObject);
var
  sInputString: string;
begin
  sInputString := InputBox('Data-Watchdog', 'Neue eMail-Adresse', '');
  if sInputString <> '' then
    lbEmail.Items.Add(sInputString);

end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.bbnDeleteClick(Sender: TObject);
var
  i: Integer;
begin

  i := lbEmail.ItemIndex;
  if i > -1 then
    lbEmail.Items.Delete(i);

end;

(* ****************************************************************** *)
function TfrmDataWatchdog.bIsAllskyDir(sFN: string): Boolean;
(* Schema 2021-06-29 *)
begin
  result := true;
  result := result and lStrIsInt(copy(sFN, 1, 4));
  result := result and lStrIsInt(copy(sFN, 6, 2));
  result := result and lStrIsInt(copy(sFN, 9, 2));
end;
(* ****************************************************************** *)
function TfrmDataWatchdog.dAllskyFNToDate(sAllskyFN: string): TDateTime;
(* wandelt das Datum im Filenamen im Schema
 * 7a210629_084500_UTC+00_0125_42.jpg
 * in ein richtigs Datum um
 *)
var
  iYear, iMonth, iDay, iHour, iMinute, iSecond: Word;

begin
  iYear := StrToIntDef(copy(sAllskyFN, 3, 2), 1900)+2000;
  iMonth := StrToIntDef(copy(sAllskyFN, 5, 2), 1);
  iDay := StrToIntDef(copy(sAllskyFN, 7, 2), 1);
  iHour := StrToIntDef(copy(sAllskyFN, 10, 2), 1);
  iMinute:= StrToIntDef(copy(sAllskyFN, 12, 2), 1);
  iSecond := StrToIntDef(copy(sAllskyFN, 15, 2), 1);
  Result := EncodeDateTime(iYear, iMonth, iDay,iHour,iMinute,iSecond,0);

end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.CheckData;

var
  i: Integer;
  s, sDir, sPrefix, sFN, sMessage: string;
  dTimeDif: TDateTime;
  sr: TSearchRec;
  slsDirList: TStringList;
begin
  slsDirList := TStringList.Create;
  sMessage := '';
  // for i := 1 to cgdFileList.RowCount do
  i := 1;
  sDir := cgdFileList.cells[1, i] + '\*.*';
  if sDir <> '' then
  begin
    if SysUtils.FindFirst(sDir, faDirectory, sr) = 0 then
    begin
      repeat
        sFN := sr.Name;
        if ((sr.Attr and faDirectory) = sr.Attr) and (sFN[1] <> '.') then
        begin
          if bIsAllskyDir(sr.Name) then
          begin
            slsDirList.Add(sFN);
          end;
        end;
      until FindNext(sr) <> 0;
      FindClose(sr);
    end;

  end;
  slsDirList.Sort;
 // showmessage(slsDirList[slsDirList.Count - 1]);
  sDir := cgdFileList.cells[1, i] + '\' + slsDirList[slsDirList.Count - 1];
  // lette Datei Suchen

  if SysUtils.FindFirst(sDir+'\*.jpg', faAnyFile, sr) = 0 then
  begin
    repeat
      sFN := sr.Name;
    until FindNext(sr) <> 0;
    FindClose(sr);
  end;
 // Showmessage (sFN );       // '7a210629_084500_UTC+00_0125_42.jpg'
  if Now - dAllskyFNToDate(sFN)>1/24 then
   showmessage('Zu alt');

  // s := sDir + '\' + cgdFileList.cells[2, i] + FormatDateTime('yyyymmdd',
  // now) + '.' + cgdFileList.cells[3, i];
  // dTimeDif := dCheckLastDate(s);
  // cgdFileList.cells[4, i] := FormatDateTime('hh:mm', dTimeDif);
  // dTimeDif := frac(now) - dTimeDif;
  // cgdFileList.cells[5, i] := FormatDateTime('hh:mm', dTimeDif);
  // if dTimeDif < speTimeout.Value / (24 * 60) then
  // begin // Zeitdifferenz mehr als 1 h und mehr als 15 min
  // cgdFileList.RowColor[i] := clLime;
  // end
  // else if frac(now) > speTimeout.Value / (24 * 60) then
  // begin
  // cgdFileList.RowColor[i] := clRed;
  // sMessage := sMessage + '  Ausfall ' + cgdFileList.cells[2, i] + ', ';
  //
  // end;
  cgdFileList.Repaint;
  if sMessage <> '' then
  begin // Ausfall !!
    FailedMessage(sMessage);
    lWriteToErrorLog(sGlobLogFile, sMessage);
  end
  else
  begin
    if _bMailSent then
    begin
      SendMailToList('Wieder OK');
      lWriteToErrorLog(sGlobLogFile, 'Wieder OK');
    end;
    _bMailSent := false;
  end;

  slsDirList.free;
end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.chkAutostartClick(Sender: TObject);
begin
  if chkAutostart.Checked then
  begin
    NextAlarm; // CheckSpectromat;
  end
  else
  begin
    HSMultiAlarmClock1.aAlarmEnabled[1] := false;
  end;

end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.NextAlarm;
var
  dTime: TDateTime;
begin
  dTime := now + speInterval.Value / (24 * 60);
  StatusBar1.Panels[0].text := 'NextAlarm: ' +
    FormatDateTime('dd.mm. hh:nn:ss', dTime);
  HSMultiAlarmClock1.aAlarmTime[1] := dTime;
  HSMultiAlarmClock1.aAlarmEnabled[1] := true;
end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.HSMultiAlarmClock1Alarm(Sender: TObject);
begin
  if HSMultiAlarmClock1.aAlarm[1] then
  begin // Timetable
    HSMultiAlarmClock1.aAlarmEnabled[1] := false;
    CheckData;
    NextAlarm;
  end;
end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.FormShow(Sender: TObject);
var
  fIni: TIniFile;
  i: Integer;
  s: string;
begin
  fIni := TIniFile.Create(sGlobIniFile);
  with fIni do
  begin
    speInterval.Value := ReadInteger('Global', 'Timeout', 30);
    chkAutostart.Checked := Readbool('Autostart', 'Enabled', false);
    Application.ProcessMessages;
    chkEmail.Checked := Readbool('Global', 'SendMail', false);
    free;
  end;

end;

(* ****************************************************************** *)

function TfrmDataWatchdog.dCheckLastDate(sFileName: string): TDateTime;
var
  s, sLine: string;
  dLastEntry: TDateTime;

begin
  result := Nulldate;
  if not FileExists(sFileName) then
  begin
    result := -1;
    exit;
  end;
  assignFile(f, sFileName);
{$I-}
  Reset(f);
  while not eof(f) do
  begin
    Readln(f, sLine);
  end;
  CloseFile(f);
{$I+}
  s := sToken(sLine, [';'], true);
  sToken(s, [' '], true);
  result := StrToTime(s);

  (*
    lblScanRunning.Caption := 'Spectromat Scanning';
    lblScanRunning.Font.Color := clLime;

    sDir := deDataDir.Text;
    if rbDailySubDir.Checked then
    sDir := sDir + '\' + formatDatetime('yyyymmdd', now);
    if rbNSubDir.checked then begin
    s := sNameOfLastDir(sDir, '');
    sDir := sDir + '\' + s;
    end;

    dLastFile := dDateofLastFile(sDir, 'txt');
    lblLastFile.Caption := 'Last Filedate:' + FormatDateTime('dd.mm.yy hh:nn', dLastFile);
    dNow := Now;
    if dNow - dLastFile > TimeOut / (24 * 60) then begin
    lWriteToErrorLog(sGlobLogFile, 'File too old (' + FormatDateTime('dd.mm.yy hh:nn', dLastFile) + ')');
    MailMessage('File too old (' + FormatDateTime('dd.mm.yy hh:nn', dLastFile) + ')' + ' - Restart Spectromat');
    KillTask(ExtractFileName(feSpectromatExe.Text));
    RestartSpectromatAndSkyscanner;
    end
    else // dann hat also alles geklappt
    _bMailSent := false;
    end
    else begin
    lblScanRunning.Caption := 'Spectromat not Scanning';
    lblScanRunning.Font.Color := clBlack;
    end;
  *)

end;
(* ****************************************************************** *)

procedure TfrmDataWatchdog.FailedMessage(sMessage: string);
var
  k: Integer;

begin
  if not _bMailSent then
    SendMailToList(sMessage);
  (* for k := 0 to lbEmail.Items.Count - 1 do
    try
    IdSMTP1.QuickSend('mailgate.uni-hannover.de', sMessage, lbEmail.items[k], 'schilke@muk.uni-hannover.de', sMessage);
    except
    lWriteToErrorLog(sGlobLogFile, ' Email an "' + lbEmail.items[k] + '" konnte nicht versandt werden');
    end; *)
  _bMailSent := true;

end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.bbnAddPathClick(Sender: TObject);
begin
  cgdFileList.cells[0, cgdFileList.RowCount - 1] :=
    inttostr(cgdFileList.RowCount - 1);
  cgdFileList.cells[1, cgdFileList.RowCount - 1] := deAllskyPath.Directory;
  cgdFileList.cells[2, cgdFileList.RowCount - 1] := slePrefix.text;
  cgdFileList.cells[3, cgdFileList.RowCount - 1] := sleExt.text;
  cgdFileList.RowCount := cgdFileList.RowCount + 1;
end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.feDataFileAfterDialog(Sender: TObject;
  var Name: string; var Action: Boolean);
begin
  if Action then
  begin
    AnalyseFileName(name, _sPath, _sFN, _sExt);

    slePrefix.text := copy(_sFN, 1, length(_sFN) - 8);
    sleExt.text := _sExt;
  end;
end;

(* ****************************************************************** *)

procedure TfrmDataWatchdog.bbnCheckClick(Sender: TObject);
begin
  CheckData;

end;
(* ****************************************************************** *)

procedure TfrmDataWatchdog.btnTestPowermanagerClick(Sender: TObject);
begin
  ReconnectLogger;
end;
(* ****************************************************************** *)

procedure TfrmDataWatchdog.ReconnectLogger;
var
  sMessage, sParams: string;
begin
  //
  // if not bShellExecute(fePowerManager.Text, '', sMessage) then
  // showError(sMessage)
  // else begin
  // sParams := '-off -' + sleDevice.Text + '  -' + sleSocket1.Text;
  // lWriteToErrorLog(sGlobLogFile, 'Switch OFF Logger');
  // bShellExecute(fePowerManager.Text, sParams, sMessage);
  // sleep(10000); // 10s solten reichen
  // sParams := '-on -' + sleDevice.Text + '  -' + sleSocket1.Text;
  // bShellExecute(fePowerManager.Text, sParams, sMessage);
  // lWriteToErrorLog(sGlobLogFile, 'Switch ON Logger');
  // end;
end;

(* ****************************************************************** *)

initialization

(* Es soll nur eine Instanz laufen *)

Mutex := CreateMutex(nil, true, 'DataWatchdogMutex');
if getLastError = ERROR_ALREADY_EXISTS then
begin
  h := 0;
  repeat
    h := FindWindowEx(0, h, 'TFrmDataWatchdog', PCHAR('Data-Watchdog'));
  until h <> Application.handle;
  if h <> 0 then
  begin
    Windows.showWindow(h, SW_ShowNormal);
    Windows.SetForegroundWindow(h);
  end;
  halt;
end;

finalization

ReleaseMutex(Mutex);

end.
