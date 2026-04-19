#define AppVersion GetFileVersion("..\bin\Release\publish\EasyComServer.exe")

[Setup]
AppName=EasyComServer
AppVersion={#AppVersion}
AppPublisher=Michael Fritzsche
DefaultDirName={autopf}\EasyComServer
DefaultGroupName=EasyComServer
OutputDir=.
OutputBaseFilename=EasyComServer_Setup_{#AppVersion}
Compression=lzma
SolidCompression=yes
SetupIconFile=..\easycom.ico
PrivilegesRequired=admin

[Files]
; EASY_COM.dll is excluded from the package — redistribution is not permitted.
; The installer downloads it at runtime from the Eaton server.
Source: "..\bin\Release\publish\*"; DestDir: "{app}"; Flags: recursesubdirs; Excludes: "EASY_COM.dll,easycom.ini"
; easycom.ini: only written if not already present; never removed by the uninstaller
; (we handle deletion ourselves in CurUninstallStepChanged after asking the user)
Source: "..\bin\Release\publish\easycom.ini"; DestDir: "{app}"; Flags: uninsneveruninstall onlyifdoesntexist

[Run]
Filename: "sc"; Parameters: "create EasyComServer binPath=""{app}\EasyComServer.exe"" start=auto obj=""LocalSystem"" DisplayName=""Moeller EASY COM Server"""; Flags: runhidden
Filename: "sc"; Parameters: "description EasyComServer ""HTTP and Telnet gateway for EASY_COM.dll"""; Flags: runhidden
Filename: "sc"; Parameters: "start EasyComServer"; Flags: runhidden

[UninstallRun]
Filename: "sc"; Parameters: "stop EasyComServer"; Flags: runhidden; RunOnceId: "StopService"
Filename: "cmd"; Parameters: "/c timeout /t 3 /nobreak"; Flags: runhidden; RunOnceId: "Wait"
Filename: "sc"; Parameters: "delete EasyComServer"; Flags: runhidden; RunOnceId: "DeleteService"

[Icons]
Name: "{group}\Uninstall"; Filename: "{uninstallexe}"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "german";  MessagesFile: "compiler:Languages\German.isl"

[CustomMessages]
; ── Status labels shown during DLL download ───────────────────────────────────
english.DllProbing=Checking for latest EASY_COM version on Eaton server ...
english.DllDownloading=Downloading %1 from Eaton ...
english.DllExtracting=Extracting %1 ...

german.DllProbing=Suche aktuellste EASY_COM-Version auf Eaton-Server ...
german.DllDownloading=Lade %1 von Eaton herunter ...
german.DllExtracting=Entpacke %1 ...

; ── Readme offer ──────────────────────────────────────────────────────────────
english.ReadmeFound=Documentation found: %1
english.ReadmeOpen=Open now?

german.ReadmeFound=Dokumentation gefunden: %1
german.ReadmeOpen=Jetzt anzeigen?

; ── Manual DLL selection ──────────────────────────────────────────────────────
english.DllBrowsePrompt=Select EASY_COM.dll from disk?%nThe file will be copied to the installation directory.
english.DllBrowseTitle=Select EASY_COM.dll
english.DllBrowseFilter=EASY_COM.dll|EASY_COM.dll|DLL files (*.dll)|*.dll
english.DllCopyError=Error copying file:

german.DllBrowsePrompt=EASY_COM.dll manuell von der Festplatte auswählen?%nDie Datei wird in das Installationsverzeichnis kopiert.
german.DllBrowseTitle=EASY_COM.dll auswählen
german.DllBrowseFilter=EASY_COM.dll|EASY_COM.dll|DLL-Dateien (*.dll)|*.dll
german.DllCopyError=Fehler beim Kopieren der Datei:

; ── Wizard pages ──────────────────────────────────────────────────────────────
english.PagePortTitle=Port and Interface Configuration
english.PagePortSubtitle=Configure the ports and the COM interface.
english.PagePortHttp=HTTP port (web console):
english.PagePortTelnet=Telnet port:
english.PagePortCom=COM port (1=COM1, 2=COM2 ...):
english.PagePortBaud=Baud rate:
english.PageAuthTitle=Access Protection (optional)
english.PageAuthSubtitle=HTTP Basic Authentication for the web console.
english.PageAuthDesc=Leave blank to disable authentication.
english.PageAuthUser=Username:
english.PageAuthPass=Password:

german.PagePortTitle=Port- und Schnittstellenkonfiguration
german.PagePortSubtitle=Ports und COM-Schnittstelle konfigurieren.
german.PagePortHttp=HTTP-Port (Webkonsole):
german.PagePortTelnet=Telnet-Port:
german.PagePortCom=COM-Port (1=COM1, 2=COM2 ...):
german.PagePortBaud=Baudrate:
german.PageAuthTitle=Zugriffsschutz (optional)
german.PageAuthSubtitle=HTTP-Basic-Authentifizierung für die Webkonsole.
german.PageAuthDesc=Leer lassen, um die Authentifizierung zu deaktivieren.
german.PageAuthUser=Benutzername:
german.PageAuthPass=Passwort:

; ── Validation ────────────────────────────────────────────────────────────────
english.ErrHttpPort=Please enter a valid HTTP port (1-65535).
english.ErrTelnetPort=Please enter a valid Telnet port (1-65535).
english.ErrPortsEqual=HTTP port and Telnet port must be different.
english.ErrComPort=Please enter a valid COM port number (1-32).
english.ErrBaudRate=Please enter a valid baud rate (300-115200).
english.WarnNoPassword=No password specified.%nWithout a password the web console is unprotected.%nContinue anyway?

german.ErrHttpPort=Bitte einen gültigen HTTP-Port eingeben (1–65535).
german.ErrTelnetPort=Bitte einen gültigen Telnet-Port eingeben (1–65535).
german.ErrPortsEqual=HTTP-Port und Telnet-Port müssen unterschiedlich sein.
german.ErrComPort=Bitte eine gültige COM-Port-Nummer eingeben (1–32).
german.ErrBaudRate=Bitte eine gültige Baudrate eingeben (300–115200).
german.WarnNoPassword=Kein Passwort angegeben.%nOhne Passwort ist die Webkonsole ungeschützt.%nTrotzdem fortfahren?

; ── Existing config detection ─────────────────────────────────────────────────
english.ExistingConfigFound=An existing configuration file was found.%nPort and authentication settings from the previous installation will be kept.%nThe configuration pages will be skipped.
german.ExistingConfigFound=Eine vorhandene Konfigurationsdatei wurde gefunden.%nPort- und Authentifizierungseinstellungen der vorherigen Installation werden übernommen.%nDie Konfigurationsseiten werden übersprungen.

; ── Uninstall ─────────────────────────────────────────────────────────────────
english.UninstallKeepConfig=Keep the configuration file (easycom.ini)?%n%nYes — settings are preserved for a future reinstallation.%nNo  — the file is deleted.
german.UninstallKeepConfig=Konfigurationsdatei (easycom.ini) behalten?%n%nJa  — Einstellungen bleiben für eine spätere Neuinstallation erhalten.%nNein — die Datei wird gelöscht.

; ── Download failed messages ──────────────────────────────────────────────────
english.DllDownloadFailed=EASY_COM.dll could not be downloaded automatically.%n(No internet access or Eaton server unreachable.)
english.DllNotInstalled=EASY_COM.dll was not installed.
english.DllManualHint=Please download manually and copy EASY_COM.dll to the installation directory:
english.DllDownloadUrl=https://es-assets.eaton.com/AUTOMATION/DOWNLOAD/DOWNLOADCENTER/EASY/LIB/EASY_COM_V250.zip

german.DllDownloadFailed=EASY_COM.dll konnte nicht automatisch heruntergeladen werden.%n(Kein Internetzugang oder Eaton-Server nicht erreichbar.)
german.DllNotInstalled=EASY_COM.dll wurde nicht installiert.
german.DllManualHint=Bitte manuell herunterladen und EASY_COM.dll in das Installationsverzeichnis kopieren:
german.DllDownloadUrl=https://es-assets.eaton.com/AUTOMATION/DOWNLOAD/DOWNLOADCENTER/EASY/LIB/EASY_COM_V250.zip

[Code]
var
  PortPage: TInputQueryWizardPage;
  AuthPage: TInputQueryWizardPage;
  EasyDllResultFile: String;  // set by DownloadEasyComDll, read in CurStepChanged
  ExistingIniFound: Boolean;  // True when easycom.ini already exists in {app}
  KeepConfig: Boolean;        // True when user wants to keep easycom.ini on uninstall

// Compute SHA-256 via PowerShell and return "sha256:<hex>".
// Falls back to the plaintext value if PowerShell fails.
function HashPasswordSha256(const Password: String): String;
var
  PS1File, HashFile, EscapedPw: String;
  Lines: TArrayOfString;
  ResultCode, I: Integer;
  C: Char;
begin
  Result := Password;

  // Escape single quotes for a PowerShell single-quoted string: ' -> ''
  EscapedPw := '';
  for I := 1 to Length(Password) do
  begin
    C := Password[I];
    if C = '''' then
      EscapedPw := EscapedPw + ''''''
    else
      EscapedPw := EscapedPw + C;
  end;

  PS1File  := ExpandConstant('{tmp}\echash.ps1');
  HashFile := ExpandConstant('{tmp}\echash.tmp');

  SaveStringToFile(PS1File,
    '$bytes = [System.Text.Encoding]::UTF8.GetBytes(''' + EscapedPw + ''')' + #13#10 +
    '$hash  = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)' + #13#10 +
    '$hex   = ([System.BitConverter]::ToString($hash)).Replace(''-'', '''').ToLower()' + #13#10 +
    'Set-Content -Path "' + HashFile + '" -Value ("sha256:"+$hex) -Encoding ASCII -NoNewline',
    False);

  Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -File "' + PS1File + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  DeleteFile(PS1File);

  if LoadStringsFromFile(HashFile, Lines) and (GetArrayLength(Lines) > 0) then
  begin
    Result := Lines[0];
    DeleteFile(HashFile);
  end;
end;

function DotNetInstalled: Boolean;
var
  Key: String;
  Names: TArrayOfString;
  I: Integer;
begin
  Result := False;
  Key := 'SOFTWARE\dotnet\Setup\InstalledVersions\x86\sharedfx\Microsoft.NETCore.App';
  if RegGetValueNames(HKLM, Key, Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if Pos('8.', Names[I]) = 1 then
      begin
        Result := True;
        Exit;
      end;
end;

function DownloadAndInstallDotNet: Boolean;
var
  Url, TempFile: String;
  ResultCode: Integer;
begin
  Result := False;
  Url      := 'https://builds.dotnet.microsoft.com/dotnet/Runtime/8.0.25/dotnet-runtime-8.0.25-win-x86.exe';
  TempFile := ExpandConstant('{app}\dotnet8.exe');

  ForceDirectories(ExpandConstant('{app}'));
  WizardForm.StatusLabel.Caption := 'Downloading .NET 8 Runtime...';

  Exec(ExpandConstant('{cmd}'),
    '/c certutil -urlcache -split -f "' + Url + '" "' + TempFile + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if not FileExists(TempFile) then
  begin
    Exec('powershell.exe',
      '-NoProfile -ExecutionPolicy Bypass -Command "' +
      '[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;' +
      '(New-Object Net.WebClient).DownloadFile(''' + Url + ''',''' + TempFile + ''')"',
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;

  if not FileExists(TempFile) then
  begin
    MsgBox('Download failed.' + #13#10 +
           'Please install .NET 8 Runtime (x86) manually:' + #13#10 +
           'https://aka.ms/dotnet/8.0/dotnet-runtime-win-x86.exe',
           mbError, MB_OK);
    Exit;
  end;

  WizardForm.StatusLabel.Caption := 'Installing .NET 8 Runtime...';
  if Exec(TempFile, '/quiet /norestart', '', SW_HIDE,
          ewWaitUntilTerminated, ResultCode) then
  begin
    DeleteFile(TempFile);
    if (ResultCode = 0) or (ResultCode = 1638) then
      Result := True
    else
      MsgBox('.NET installation failed (code: ' +
             IntToStr(ResultCode) + ')', mbError, MB_OK);
  end
  else
    MsgBox('Could not launch the .NET installer.', mbError, MB_OK);
end;

// ─── EASY_COM DLL download ──────────────────────────────────────────────────
//
// Three-stage download with individual status labels so the installer never
// appears frozen.  Result file (EasyDllResultFile) contains 3 lines:
//   [0] full DLL path
//   [1] relative path with forward slashes  (written to dll_path in ini)
//   [2] extraction directory               (searched for readme files)
//
// Stage 1 – version probe  : HEAD V251..V260, 3 s timeout, falls back to V250
// Stage 2 – ZIP download   : Net.WebClient.DownloadFile
// Stage 3 – extract + find : Expand-Archive, locate EASY_COM.dll recursively
//
function DownloadEasyComDll(const AppDir: String): Boolean;
var
  PS1, VerFile, Script, ZipName, ZipPath, ExtractDir: String;
  Lines: TArrayOfString;
  BestVer, BestUrl: String;
  ResultCode: Integer;
begin
  Result := False;
  PS1               := ExpandConstant('{tmp}\ecps.ps1');
  VerFile           := ExpandConstant('{tmp}\ecver.txt');
  EasyDllResultFile := ExpandConstant('{tmp}\ecdll.txt');
  DeleteFile(VerFile);
  DeleteFile(EasyDllResultFile);

  // ── Stage 1: probe for latest version ────────────────────────────────────
  WizardForm.StatusLabel.Caption := CustomMessage('DllProbing');
  WizardForm.StatusLabel.Update;

  Script :=
    'param([string]$Out)' + #13#10 +
    '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12' + #13#10 +
    '$base = ' + #39 + 'https://es-assets.eaton.com/AUTOMATION/DOWNLOAD/DOWNLOADCENTER/EASY/LIB' + #39 + #13#10 +
    '$ver = 250; $url = "$base/EASY_COM_V250.zip"' + #13#10 +
    'for ($v = 251; $v -le 260; $v++) {' + #13#10 +
    '    try {' + #13#10 +
    '        $r = [Net.HttpWebRequest]::Create("$base/EASY_COM_V$v.zip")' + #13#10 +
    '        $r.Method = ' + #39 + 'HEAD' + #39 + '; $r.Timeout = 3000' + #13#10 +
    '        ($r.GetResponse()).Close()' + #13#10 +
    '        $ver = $v; $url = "$base/EASY_COM_V$v.zip"' + #13#10 +
    '    } catch {}' + #13#10 +
    '}' + #13#10 +
    'Set-Content $Out "$ver|$url" -Encoding ASCII';

  SaveStringToFile(PS1, Script, False);
  Exec('powershell.exe',
       '-NoProfile -ExecutionPolicy Bypass -File "' + PS1 + '" -Out "' + VerFile + '"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  DeleteFile(PS1);

  if not LoadStringsFromFile(VerFile, Lines) or (GetArrayLength(Lines) = 0) then Exit;
  DeleteFile(VerFile);
  BestVer := Copy(Lines[0], 1, Pos('|', Lines[0]) - 1);
  BestUrl := Copy(Lines[0], Pos('|', Lines[0]) + 1, MaxInt);
  if BestVer = '' then Exit;

  // ── Stage 2: download ZIP ─────────────────────────────────────────────────
  ZipName := 'EASY_COM_V' + BestVer + '.zip';
  ZipPath := ExpandConstant('{tmp}\') + ZipName;
  WizardForm.StatusLabel.Caption := FmtMessage(CustomMessage('DllDownloading'), [ZipName]);
  WizardForm.StatusLabel.Update;

  Script :=
    'param([string]$Url, [string]$Dest)' + #13#10 +
    '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12' + #13#10 +
    '(New-Object Net.WebClient).DownloadFile($Url, $Dest)';

  SaveStringToFile(PS1, Script, False);
  Exec('powershell.exe',
       '-NoProfile -ExecutionPolicy Bypass -File "' + PS1 + '" -Url "' + BestUrl + '" -Dest "' + ZipPath + '"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  DeleteFile(PS1);
  if not FileExists(ZipPath) then Exit;

  // ── Stage 3: extract and locate EASY_COM.dll ──────────────────────────────
  ExtractDir := AppDir + '\EASY_COM_V' + BestVer;
  WizardForm.StatusLabel.Caption := FmtMessage(CustomMessage('DllExtracting'), [ZipName]);
  WizardForm.StatusLabel.Update;

  Script :=
    'param([string]$Zip, [string]$Dir, [string]$AppDir, [string]$Out)' + #13#10 +
    'if (Test-Path $Dir) { Remove-Item $Dir -Recurse -Force }' + #13#10 +
    'Expand-Archive -Path $Zip -DestinationPath $Dir -Force' + #13#10 +
    'Remove-Item $Zip -Force -ErrorAction SilentlyContinue' + #13#10 +
    '$dll = Get-ChildItem -Path $Dir -Filter ' + #39 + 'EASY_COM.dll' + #39 + ' -Recurse | Select-Object -First 1' + #13#10 +
    'if ($dll) {' + #13#10 +
    '        $rel = $dll.FullName.Substring($AppDir.TrimEnd(' + #39 + '\' + #39 + ').Length)' +
                   '.TrimStart(' + #39 + '\' + #39 + ') -replace ' + #39 + '\\' + #39 + ', ' + #39 + '/' + #39 + #13#10 +
    '    [IO.File]::WriteAllLines($Out, @($dll.FullName, $rel, $Dir))' + #13#10 +
    '} else {' + #13#10 +
    '    Set-Content $Out ' + #39 + 'ERROR:NOT_FOUND' + #39 + ' -Encoding ASCII' + #13#10 +
    '}';

  SaveStringToFile(PS1, Script, False);
  Exec('powershell.exe',
       '-NoProfile -ExecutionPolicy Bypass -File "' + PS1 + '"' +
       ' -Zip "' + ZipPath + '" -Dir "' + ExtractDir + '" -AppDir "' + AppDir + '" -Out "' + EasyDllResultFile + '"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  DeleteFile(PS1);

  if not LoadStringsFromFile(EasyDllResultFile, Lines) then Exit;
  if (GetArrayLength(Lines) = 0) or (Pos('ERROR:', Lines[0]) = 1) then Exit;

  Result := True;
end;

// Searches ExtractDir (and one subdirectory level) for readme / liesmich files.
// Prefers the file matching the active setup language; falls back to the other.
// Opens matching files with the system viewer after user confirmation.
procedure OfferReadme(const ExtractDir: String);
var
  FindRec, Sub: TFindRec;
  Pref, Fall: TArrayOfString;
  PrefCnt, FallCnt, ErrCode, I: Integer;
  Name, Primary, Secondary, List, SubDir: String;
  Paths: TArrayOfString;
  Count: Integer;
begin
  if ActiveLanguage = 'german' then
  begin Primary := 'liesmich'; Secondary := 'readme'; end
  else
  begin Primary := 'readme'; Secondary := 'liesmich'; end;

  PrefCnt := 0; FallCnt := 0;
  SetArrayLength(Pref, 10);
  SetArrayLength(Fall, 10);

  // ── Inline scan helper (repeated for top-level and each subdir) ──────────
  // Top-level .txt files
  if FindFirst(ExtractDir + '\*.txt', Sub) then
  begin
    try
      repeat
        Name := LowerCase(Sub.Name);
        if (Pos(Primary, Name) = 1) and (PrefCnt < 10) then
        begin Pref[PrefCnt] := ExtractDir + '\' + Sub.Name; PrefCnt := PrefCnt + 1; end
        else if (Pos(Secondary, Name) = 1) and (FallCnt < 10) then
        begin Fall[FallCnt] := ExtractDir + '\' + Sub.Name; FallCnt := FallCnt + 1; end;
      until not FindNext(Sub);
    finally
      FindClose(Sub);
    end;
  end;

  // One level of subdirectories ($10 = FILE_ATTRIBUTE_DIRECTORY)
  if FindFirst(ExtractDir + '\*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Attributes and $10 <> 0) and
           (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          SubDir := ExtractDir + '\' + FindRec.Name;
          if FindFirst(SubDir + '\*.txt', Sub) then
          begin
            try
              repeat
                Name := LowerCase(Sub.Name);
                if (Pos(Primary, Name) = 1) and (PrefCnt < 10) then
                begin Pref[PrefCnt] := SubDir + '\' + Sub.Name; PrefCnt := PrefCnt + 1; end
                else if (Pos(Secondary, Name) = 1) and (FallCnt < 10) then
                begin Fall[FallCnt] := SubDir + '\' + Sub.Name; FallCnt := FallCnt + 1; end;
              until not FindNext(Sub);
            finally
              FindClose(Sub);
            end;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;

  // Use preferred language; fall back to secondary if nothing found
  if PrefCnt > 0 then
  begin Paths := Pref; Count := PrefCnt; end
  else if FallCnt > 0 then
  begin Paths := Fall; Count := FallCnt; end
  else
    Exit;

  List := '';
  for I := 0 to Count - 1 do
  begin
    if List <> '' then List := List + ', ';
    List := List + ExtractFileName(Paths[I]);
  end;

  if MsgBox(FmtMessage(CustomMessage('ReadmeFound'), [List]) + #13#10#13#10 +
            CustomMessage('ReadmeOpen'),
            mbConfirmation, MB_YESNO) = IDYES then
    for I := 0 to Count - 1 do
      ShellExec('open', Paths[I], '', '', SW_SHOW, ewNoWait, ErrCode);
end;

// ────────────────────────────────────────────────────────────────────────────

// Zeigt einen Windows-Datei-öffnen-Dialog (über PowerShell + WinForms),
// kopiert die gewählte EASY_COM.dll in {app} und trägt sie in easycom.ini ein.
// Gibt True zurück, wenn die Datei erfolgreich kopiert wurde.
function BrowseForEasyComDll(const AppDir: String): Boolean;
var
  PS1, ResultFile, Script, SrcPath, DestPath, IniFile: String;
  ResultCode: Integer;
  Lines: TArrayOfString;
begin
  Result := False;

  if MsgBox(CustomMessage('DllBrowsePrompt'), mbConfirmation, MB_YESNO) <> IDYES then Exit;

  PS1        := ExpandConstant('{tmp}\browsedll.ps1');
  ResultFile := ExpandConstant('{tmp}\browsedll_result.txt');
  DeleteFile(ResultFile);

  Script :=
    'param([string]$Title, [string]$Filter, [string]$Out)' + #13#10 +
    'Add-Type -AssemblyName System.Windows.Forms' + #13#10 +
    '$dlg = New-Object System.Windows.Forms.OpenFileDialog' + #13#10 +
    '$dlg.Title    = $Title' + #13#10 +
    '$dlg.Filter   = $Filter' + #13#10 +
    '$dlg.FileName = ' + #39 + 'EASY_COM.dll' + #39 + #13#10 +
    'if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {' + #13#10 +
    '    Set-Content $Out $dlg.FileName -Encoding ASCII' + #13#10 +
    '}';

  SaveStringToFile(PS1, Script, False);

  // -STA required for WinForms dialogs; SW_HIDE suppresses the console window
  Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -STA -File "' + PS1 + '"' +
    ' -Title "' + CustomMessage('DllBrowseTitle') + '"' +
    ' -Filter "' + CustomMessage('DllBrowseFilter') + '"' +
    ' -Out "' + ResultFile + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  DeleteFile(PS1);

  if not LoadStringsFromFile(ResultFile, Lines) then Exit;
  if GetArrayLength(Lines) = 0 then Exit;

  SrcPath := Trim(Lines[0]);
  DeleteFile(ResultFile);

  if not FileExists(SrcPath) then Exit;

  DestPath := AppDir + '\EASY_COM.dll';

  if not CopyFile(SrcPath, DestPath, False) then
  begin
    MsgBox(CustomMessage('DllCopyError') + #13#10 + SrcPath, mbError, MB_OK);
    Exit;
  end;

  IniFile := AppDir + '\easycom.ini';
  SetIniString('global', 'dll_path', 'EASY_COM.dll', IniFile);

  Result := True;
end;

// ────────────────────────────────────────────────────────────────────────────

procedure InitializeWizard;
var
  IniFile: String;
begin
  PortPage := CreateInputQueryPage(wpSelectDir,
    CustomMessage('PagePortTitle'),
    CustomMessage('PagePortSubtitle'),
    '');
  PortPage.Add(CustomMessage('PagePortHttp'),   False);
  PortPage.Add(CustomMessage('PagePortTelnet'), False);
  PortPage.Add(CustomMessage('PagePortCom'),    False);
  PortPage.Add(CustomMessage('PagePortBaud'),   False);

  AuthPage := CreateInputQueryPage(PortPage.ID,
    CustomMessage('PageAuthTitle'),
    CustomMessage('PageAuthSubtitle'),
    CustomMessage('PageAuthDesc'));
  AuthPage.Add(CustomMessage('PageAuthUser'), False);
  AuthPage.Add(CustomMessage('PageAuthPass'), True);

  // Check for an existing easycom.ini from a previous installation
  IniFile := ExpandConstant('{app}\easycom.ini');
  ExistingIniFound := FileExists(IniFile);

  if ExistingIniFound then
  begin
    // Pre-populate wizard values from existing ini so the user sees what is kept
    PortPage.Values[0] := GetIniString('instance', 'http_port',   '8083', IniFile);
    PortPage.Values[1] := GetIniString('instance', 'telnet_port', '23',   IniFile);
    PortPage.Values[2] := GetIniString('instance', 'com_port',    '1',    IniFile);
    PortPage.Values[3] := GetIniString('instance', 'baud_rate',   '9600', IniFile);
    AuthPage.Values[0] := GetIniString('instance', 'auth_user',   'admin', IniFile);
    AuthPage.Values[1] := '';  // never pre-fill password field

    MsgBox(CustomMessage('ExistingConfigFound'), mbInformation, MB_OK);
  end
  else
  begin
    // Default values for a fresh installation
    PortPage.Values[0] := '8083';
    PortPage.Values[1] := '23';
    PortPage.Values[2] := '1';
    PortPage.Values[3] := '9600';
    AuthPage.Values[0] := 'admin';
    AuthPage.Values[1] := '';
  end;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  // Skip port and auth pages when an existing configuration is being kept
  Result := ExistingIniFound and
            ((PageID = PortPage.ID) or (PageID = AuthPage.ID));
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Http, Telnet, Com, Baud: Integer;
begin
  Result := True;

  if CurPageID = PortPage.ID then
  begin
    Http   := StrToIntDef(PortPage.Values[0], 0);
    Telnet := StrToIntDef(PortPage.Values[1], 0);
    Com    := StrToIntDef(PortPage.Values[2], 0);
    Baud   := StrToIntDef(PortPage.Values[3], 0);

    if (Http < 1) or (Http > 65535) then
    begin
      MsgBox(CustomMessage('ErrHttpPort'), mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Telnet < 1) or (Telnet > 65535) then
    begin
      MsgBox(CustomMessage('ErrTelnetPort'), mbError, MB_OK);
      Result := False; Exit;
    end;
    if Http = Telnet then
    begin
      MsgBox(CustomMessage('ErrPortsEqual'), mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Com < 1) or (Com > 32) then
    begin
      MsgBox(CustomMessage('ErrComPort'), mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Baud < 300) or (Baud > 115200) then
    begin
      MsgBox(CustomMessage('ErrBaudRate'), mbError, MB_OK);
      Result := False; Exit;
    end;
  end;

  if CurPageID = AuthPage.ID then
  begin
    if (AuthPage.Values[0] <> '') and (AuthPage.Values[1] = '') then
    begin
      if MsgBox(CustomMessage('WarnNoPassword'), mbConfirmation, MB_YESNO) = IDNO then
      begin
        Result := False; Exit;
      end;
    end;
  end;

  if CurPageID = wpReady then
  begin
    if not DotNetInstalled then
    begin
      if MsgBox('.NET 8 Runtime (x86) is not installed.' + #13#10 +
                'Download and install it now? (~26 MB)',
                mbConfirmation, MB_YESNO) = IDYES then
      begin
        if not DownloadAndInstallDotNet then
        begin
          Result := False; Exit;
        end;
      end
      else
      begin
        MsgBox('EasyComServer requires .NET 8 Runtime (x86).' + #13#10 +
               'Setup will be cancelled.', mbError, MB_OK);
        Result := False; Exit;
      end;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  IniFile, HtmlFile: String;
  HttpPort, TelnetPort, ComPort, BaudRate: String;
  AuthUser, AuthPass: String;
  ShellDir, ShellLink: String;
  DllLines: TArrayOfString;
  Dummy: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    IniFile := ExpandConstant('{app}\easycom.ini');

    if not ExistingIniFound then
    begin
      // Fresh installation: write wizard values to easycom.ini
      HttpPort   := PortPage.Values[0];
      TelnetPort := PortPage.Values[1];
      ComPort    := PortPage.Values[2];
      BaudRate   := PortPage.Values[3];
      AuthUser   := AuthPage.Values[0];
      AuthPass   := AuthPage.Values[1];

      SetIniString('instance', 'http_port',   HttpPort,   IniFile);
      SetIniString('instance', 'telnet_port', TelnetPort, IniFile);
      SetIniString('instance', 'com_port',    ComPort,    IniFile);
      SetIniString('instance', 'baud_rate',   BaudRate,   IniFile);

      if (AuthUser <> '') and (AuthPass <> '') then
      begin
        SetIniString('instance', 'basic_auth', 'true',                       IniFile);
        SetIniString('instance', 'auth_user',  AuthUser,                     IniFile);
        SetIniString('instance', 'auth_pass',  HashPasswordSha256(AuthPass), IniFile);
      end
      else
        SetIniString('instance', 'basic_auth', 'false', IniFile);
    end
    else
    begin
      // Reinstall: read existing values for URL patching below
      HttpPort := GetIniString('instance', 'http_port', '8083', IniFile);
    end;

    // Patch index.html: update the default server URL via PowerShell script
    HtmlFile := ExpandConstant('{app}\wwwroot\index.html');
    if FileExists(HtmlFile) then
    begin
      SaveStringToFile(ExpandConstant('{tmp}\patch.ps1'),
        '$f = "' + HtmlFile + '"' + #13#10 +
        '$c = [IO.File]::ReadAllText($f)' + #13#10 +
        '$c = $c -replace "url: ''http://localhost:8083''", "url: ''http://localhost:' + HttpPort + '''"' + #13#10 +
        '[IO.File]::WriteAllText($f, $c)',
        False);
      Exec('powershell.exe',
        '-NoProfile -ExecutionPolicy Bypass -File "' +
        ExpandConstant('{tmp}\patch.ps1') + '"',
        '', SW_HIDE, ewWaitUntilTerminated, Dummy);
      DeleteFile(ExpandConstant('{tmp}\patch.ps1'));
    end;

    // Register HTTP URL ACL
    Exec('netsh',
      'http add urlacl url="http://+:' + HttpPort + '/" user="Everyone"',
      '', SW_HIDE, ewWaitUntilTerminated, Dummy);

    // Create Start Menu shortcut
    ShellDir  := ExpandConstant('{commonprograms}\EasyComServer');
    ShellLink := ShellDir + '\EasyComServer Web Console.url';
    ForceDirectories(ShellDir);
    SaveStringToFile(ShellLink,
      '[InternetShortcut]' + #13#10 +
      'URL=http://localhost:' + HttpPort + '/' + #13#10,
      False);

    // ── Download EASY_COM.dll from Eaton and update easycom.ini ───────────
    // Result file has 3 lines: [0] full path  [1] rel path  [2] extract dir
    if DownloadEasyComDll(ExpandConstant('{app}')) then
    begin
      if LoadStringsFromFile(EasyDllResultFile, DllLines) and
         (GetArrayLength(DllLines) >= 3) then
      begin
        SetIniString('global', 'dll_path', DllLines[1], IniFile);
        OfferReadme(DllLines[2]);
      end;
      DeleteFile(EasyDllResultFile);
    end
    else
    begin
      MsgBox(CustomMessage('DllDownloadFailed'), mbInformation, MB_OK);
      if not BrowseForEasyComDll(ExpandConstant('{app}')) then
        MsgBox(CustomMessage('DllNotInstalled') + #13#10#13#10 +
               CustomMessage('DllManualHint') + #13#10 +
               CustomMessage('DllDownloadUrl') + #13#10#13#10 +
               ExpandConstant('{app}'),
               mbInformation, MB_OK);
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ShellLink, HttpPort, IniFile: String;
  Dummy: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    // Remove HTTP URL ACL
    IniFile  := ExpandConstant('{app}\easycom.ini');
    HttpPort := GetIniString('instance', 'http_port', '8083', IniFile);
    Exec('netsh',
      'http delete urlacl url="http://+:' + HttpPort + '/"',
      '', SW_HIDE, ewWaitUntilTerminated, Dummy);

    // Ask whether to keep easycom.ini for a future reinstallation
    KeepConfig := MsgBox(CustomMessage('UninstallKeepConfig'),
                         mbConfirmation, MB_YESNO) = IDYES;
  end;

  if CurUninstallStep = usPostUninstall then
  begin
    // Remove Start Menu shortcut
    ShellLink := ExpandConstant('{commonprograms}\EasyComServer\EasyComServer Web Console.url');
    DeleteFile(ShellLink);
    RemoveDir(ExpandConstant('{commonprograms}\EasyComServer'));

    // Delete easycom.ini only when the user opted not to keep it
    if not KeepConfig then
      DeleteFile(ExpandConstant('{app}\easycom.ini'));
  end;
end;
