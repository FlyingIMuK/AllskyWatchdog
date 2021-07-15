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
  Inifiles, math, system.DateUtils, TypeEdit;

const
  TimeOut = 5; // Minute (Max. Dateialter)

type
  TfrmAllskyWatchdog = class(TForm)
    bbnClose: TBitBtn;
    speInterval: TSpinEdit;
    Panel1: TPanel;
    chkEmail: TCheckBox;
    Mail: TButton;
    lbEmail: TListBox;
    bbnAddMail: TBitBtn;
    bbnDeleteMail: TBitBtn;
    IdSMTP1: TIdSMTP;
    HSMultiAlarmClock1: THSMultiAlarmClock;
    StatusBar1: TStatusBar;
    Label2: TLabel;
    chkAutostart: TCheckBox;
    cgdFileList: TXColorGrid;
    bbnCheck: TBitBtn;
    Panel2: TPanel;
    deAllskyPath: TJvDirectoryEdit;
    Label1: TLabel;
    tedMaxDiff: TTypeEdit;
    bbnAddPath: TBitBtn;
    btnDelPath: TButton;
    Panel3: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    chkPowerManager: TCheckBox;
    sleDevice: TEdit;
    sleSocket: TEdit;
    btnTestPowermanager: TButton;
    fePowerManager: TJvFilenameEdit;
    (* procedure SmtpCli1RequestDone(Sender: TObject; RqType: TSmtpRequest;
      ErrorCode: Word); *)
    procedure FormCreate(Sender: TObject);
    procedure MailClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bbnHelpClick(Sender: TObject);
    procedure bbnCloseClick(Sender: TObject);
    procedure bbnAddMailClick(Sender: TObject);
    procedure bbnDeleteMailClick(Sender: TObject);
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
    procedure btnDelPathClick(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure cgdFileListSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure deAllskyPathChange(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ReconnectPowermanager;
  private
    _bMailSent: Boolean;
    hMainHandle, hErrorWindow: HWND;
    _sPath, _sFN, _sExt: string;
    _iRow: Integer;

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
  frmAllskyWatchdog: TfrmAllskyWatchdog;

implementation

uses fMsgWin, mGlobProc_B, mMUKSunCalc_B, mApiFunctions_B;

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

procedure TfrmAllskyWatchdog.FormCreate(Sender: TObject);
var
  fIni: TIniFile;
  sDir, sFN, sExe, sName, sExtension: string;
  i, iMaxDiff: Integer;
  dDate: TDateTime;
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
  Caption := 'Allksy-Watchdog  ' +
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

    chkAutostart.Checked := Readbool('Autostart', 'Enabled', false);

    fePowerManager.text := ReadString('PowerManager', 'Path', '');
    sleDevice.text := ReadString('PowerManager', 'Device', '');
    sleSocket.text := ReadString('Powermanager', 'Socket', '');
    chkPowerManager.Checked:=ReadBool('PowerManager', 'Used', false);

    // muss hinter speInterval stehen
    chkEmail.Checked := Readbool('Global', 'SendMail', false);
    i := 1;
    cgdFileList.ClearGrid;
    repeat

      sFN := ReadString('Files', 'Path' + inttostr(i), '');
      iMaxDiff := ReadInteger('Files', 'MaxDiff' + inttostr(i), 30);
      if sFN <> '' then
      begin
        with cgdFileList do
        begin
          cells[0, 0] := 'Nr';
          cells[1, 0] := 'Path';
          cells[2, 0] := 'Max Difference(min)';
          cells[3, 0] := 'LastTime';
          cells[4, 0] := 'Difference';

          cells[0, i] := inttostr(i);
          cells[1, i] := sFN;
          cells[2, i] := inttostr(iMaxDiff);
        end;
        cgdFileList.RowCount := cgdFileList.RowCount + 1;
      end;
      inc(i);
    until sFN = '';
    cgdFileList.RowCount := cgdFileList.RowCount - 1;

    i := 0;
    repeat
      sFN := ReadString('Adresse', 'Nr' + inttostr(i), '');
      if sFN <> '' then
        lbEmail.Items.Add(sFN);
      inc(i);
    until sFN = '';
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
  _iRow := Badindex;

end;
(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.MailClick(Sender: TObject);
begin
  SendMailToList('Allsky-Watchdog-Testmail');
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.SendMailToList(sMessage: string);
var
  i: Integer;
begin
  for i := 0 to lbEmail.Items.Count - 1 do
    IdSMTP1.QuickSend('mailgate.uni-hannover.de', sMessage, lbEmail.Items[i],
      'schilke@muk.uni-hannover.de', sMessage);

end;

procedure TfrmAllskyWatchdog.StringGrid1Click(Sender: TObject);
begin
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.FormClose(Sender: TObject;
  var Action: TCloseAction);
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
        WriteInteger('Files', 'MaxDiff' + inttostr(i),
          StrToInt(cgdFileList.cells[2, i]));
      end;

    WriteBool('Autostart', 'Enabled', chkAutostart.Checked);
    WriteInteger('Global', 'Interval', speInterval.Value);

    WriteBool('Global', 'SendMail', chkEmail.Checked);
    WriteBool('PowerManager', 'Used', chkPowerManager.Checked);
    WriteString('PowerManager', 'Path', fePowerManager.text);
    WriteString('PowerManager', 'Device', sleDevice.text);
    WriteString('Powermanager', 'Socket', sleSocket.text);

    free;
  end;
  lWriteToErrorLog(sGlobLogFile, 'Application terminated');
end;
(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.bbnHelpClick(Sender: TObject);
begin
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.bbnCloseClick(Sender: TObject);
begin
  close;
end;
(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.bbnAddMailClick(Sender: TObject);
var
  sInputString: string;
begin
  sInputString := InputBox('Allsky-Watchdog', 'Neue eMail-Adresse', '');
  if sInputString <> '' then
    lbEmail.Items.Add(sInputString);

end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.bbnDeleteMailClick(Sender: TObject);
var
  i: Integer;
begin

  i := lbEmail.ItemIndex;
  if i > -1 then
    lbEmail.Items.Delete(i);

end;

(* ****************************************************************** *)
function TfrmAllskyWatchdog.bIsAllskyDir(sFN: string): Boolean;
(* Schema 2021-06-29 *)
begin
  result := true;
  result := result and lStrIsInt(copy(sFN, 1, 4));
  result := result and lStrIsInt(copy(sFN, 6, 2));
  result := result and lStrIsInt(copy(sFN, 9, 2));
end;

(* ****************************************************************** *)
function TfrmAllskyWatchdog.dAllskyFNToDate(sAllskyFN: string): TDateTime;
(* wandelt das Datum im Filenamen im Schema
  * 7a210629_084500_UTC+00_0125_42.jpg
  * in ein richtigs Datum um
*)
var
  iYear, iMonth, iDay, iHour, iMinute, iSecond: Word;

begin
  iYear := StrToIntDef(copy(sAllskyFN, 3, 2), 1900) + 2000;
  iMonth := StrToIntDef(copy(sAllskyFN, 5, 2), 1);
  iDay := StrToIntDef(copy(sAllskyFN, 7, 2), 1);
  iHour := StrToIntDef(copy(sAllskyFN, 10, 2), 1);
  iMinute := StrToIntDef(copy(sAllskyFN, 12, 2), 1);
  iSecond := StrToIntDef(copy(sAllskyFN, 15, 2), 1);
  result := EncodeDateTime(iYear, iMonth, iDay, iHour, iMinute, iSecond, 0);

end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.CheckData;

var
  i, iMaxDiffMin: Integer;
  s, sDir, sPrefix, sFN, sMessage: string;
  dLastAllsky, dTimeDif: TDateTime;
  sr: TSearchRec;
  slsDirList: TStringList;
begin
  slsDirList := TStringList.Create;
  sMessage := '';
  for i := 1 to cgdFileList.RowCount do
  begin
    slsDirList.Clear;
    sDir := cgdFileList.cells[1, i];
    if sDir <> '' then
    begin
      sDir := cgdFileList.cells[1, i] + '\*.*';
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

      slsDirList.Sort;
      // showmessage(slsDirList[slsDirList.Count - 1]);
      sDir := cgdFileList.cells[1, i] + '\' + slsDirList[slsDirList.Count - 1];
      // lette Datei Suchen

      if SysUtils.FindFirst(sDir + '\*.jpg', faAnyFile, sr) = 0 then
      begin
        repeat
          sFN := sr.Name;
        until FindNext(sr) <> 0;
        FindClose(sr);
      end;
      // Showmessage (sFN );       // '7a210629_084500_UTC+00_0125_42.jpg'
      dLastAllsky := dAllskyFNToDate(sFN);
      dTimeDif := NowUTC - dLastAllsky;
      // s := sDir + '\' + cgdFileList.cells[2, i] + FormatDateTime('yyyymmdd',
      // now) + '.' + cgdFileList.cells[3, i];
      // dTimeDif := dCheckLastDate(s);
      cgdFileList.cells[3, i] := FormatDateTime('hh:mm', dLastAllsky);
      cgdFileList.cells[4, i] := FormatDateTime('hh:mm', dTimeDif);
      iMaxDiffMin := StrToIntDef(cgdFileList.cells[2, i], 30);
      if dTimeDif < iMaxDiffMin / (24 * 60) then
      begin // Zeitdifferenz mehr als 1 h und mehr als 15 min
        cgdFileList.RowColor[i] := clLime;
      end
      else if frac(Now) > iMaxDiffMin / (24 * 60) then
      begin
        cgdFileList.RowColor[i] := clRed;
        sMessage := sMessage + '  Ausfall Allsky';
      end;
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
    end;
  end;
  slsDirList.free;
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.chkAutostartClick(Sender: TObject);
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

procedure TfrmAllskyWatchdog.NextAlarm;
var
  dTime: TDateTime;
begin
  dTime := Now + speInterval.Value / (24 * 60);
  StatusBar1.Panels[0].text := 'NextAlarm: ' +
    FormatDateTime('dd.mm. hh:nn:ss', dTime);
  HSMultiAlarmClock1.aAlarmTime[1] := dTime;
  HSMultiAlarmClock1.aAlarmEnabled[1] := true;
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.HSMultiAlarmClock1Alarm(Sender: TObject);
var
  fAzi, fZen, fBPLat, fBPLon: Extended;
begin
  if HSMultiAlarmClock1.aAlarm[1] then
  begin // Timetable
    HSMultiAlarmClock1.aAlarmEnabled[1] := false;
    Muk_zen_azi(fAzi, fZen, fBPLat, fBPLon, 52, 10, Now);
    if fZen < 80 then
      CheckData;
    NextAlarm;
  end;
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.FormShow(Sender: TObject);
var
  fIni: TIniFile;
  i: Integer;
  s: string;
begin
  fIni := TIniFile.Create(sGlobIniFile);
  with fIni do
  begin
//    speInterval.Value := ReadInteger('Global', 'Timeout', 30);
    chkAutostart.Checked := Readbool('Autostart', 'Enabled', false);
    Application.ProcessMessages;
    chkEmail.Checked := Readbool('Global', 'SendMail', false);
    free;
  end;

end;

(* ****************************************************************** *)

function TfrmAllskyWatchdog.dCheckLastDate(sFileName: string): TDateTime;
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

end;
(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.deAllskyPathChange(Sender: TObject);
begin
  bbnAddPath.Enabled := deAllskyPath.text <> '';

end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.FailedMessage(sMessage: string);
var
  k: Integer;

begin
  if not _bMailSent then
    SendMailToList(sMessage);
  _bMailSent := true;

end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.bbnAddPathClick(Sender: TObject);
begin
  cgdFileList.RowCount := cgdFileList.RowCount + 1;
  cgdFileList.cells[0, cgdFileList.RowCount - 1] :=
    inttostr(cgdFileList.RowCount - 1);
  cgdFileList.cells[1, cgdFileList.RowCount - 1] := deAllskyPath.Directory;
  cgdFileList.cells[2, cgdFileList.RowCount - 1] := tedMaxDiff.text;
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.feDataFileAfterDialog(Sender: TObject;
  var Name: string; var Action: Boolean);
begin
  if Action then
  begin
    AnalyseFileName(name, _sPath, _sFN, _sExt);

  end;
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.bbnCheckClick(Sender: TObject);
begin
  CheckData;

end;
(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.btnTestPowermanagerClick(Sender: TObject);
begin
  ReconnectPowermanager;
end;

(* ****************************************************************** *)
procedure TfrmAllskyWatchdog.ReconnectPowermanager;
var
  sMessage, sParams: string;
begin

  if not bShellExecute(fePowerManager.text, '', sMessage) then
    showError(sMessage)
  else
  begin
    sParams := '-off -' + sleDevice.text + '  -' + sleSocket.text;
    lWriteToErrorLog(sGlobLogFile, 'Switch OFF Logger');
    bShellExecute(fePowerManager.text, sParams, sMessage);
    sleep(10000); // 10s solten reichen
    sParams := '-on -' + sleDevice.text + '  -' + sleSocket.text;
    bShellExecute(fePowerManager.text, sParams, sMessage);
    lWriteToErrorLog(sGlobLogFile, 'Switch ON Logger');
  end;
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.Button2Click(Sender: TObject);
var
  fAzi, fZen, fBPLat, fBPLon: Extended;
begin
  Muk_zen_azi(fAzi, fZen, fBPLat, fBPLon, 52, 10, Now);
  showmessage(FloatToStr(fZen));
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.btnDelPathClick(Sender: TObject);
var
  i: Integer;
begin

  showmessage(inttostr(_iRow));
  cgdFileList.DeleteRow(_iRow);
  // if i > -1 then
  // cgdFileList.Items.Delete(i);

end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.cgdFileListSelectCell(Sender: TObject;
  ACol, ARow: Integer; var CanSelect: Boolean);
begin
  _iRow := ARow;
end;

(* ****************************************************************** *)

procedure TfrmAllskyWatchdog.ReconnectLogger;
var
  sMessage, sParams: string;
begin
end;

(* ****************************************************************** *)

initialization

(* Es soll nur eine Instanz laufen *)

Mutex := CreateMutex(nil, true, 'AllskyWatchdogMutex');
if getLastError = ERROR_ALREADY_EXISTS then
begin
  h := 0;
  repeat
    h := FindWindowEx(0, h, 'TFrmAllskyWatchdog', PCHAR('Allsky-Watchdog'));
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
