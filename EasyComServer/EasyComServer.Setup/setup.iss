[Setup]
AppName=EasyComServer
AppVersion=2.2.0
AppPublisher=Michael Fritzsche
DefaultDirName={autopf}\EasyComServer
DefaultGroupName=EasyComServer
OutputDir=.
OutputBaseFilename=EasyComServer_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=..\easycom.ico
PrivilegesRequired=admin

[Files]
; EASY_COM.dll is excluded from the package — redistribution is not permitted.
; The installer downloads it at runtime from the Eaton server.
Source: "..\bin\Release\publish\*"; DestDir: "{app}"; Flags: recursesubdirs; Excludes: "EASY_COM.dll"

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

[Code]
var
  PortPage: TInputQueryWizardPage;
  AuthPage: TInputQueryWizardPage;
  EasyDllResultFile: String;  // set by DownloadEasyComDll, read in CurStepChanged

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

// Probes Eaton server for versions V251..V260 (HEAD, 3 s timeout each).
// Falls back to the static V250 URL if no newer release is found.
// Extracts the ZIP into {app}\EASY_COM_V{n} and writes four lines to
// EasyDllResultFile: version | dllFullPath | relPath (fwd-slashes) | extractDir.
// Returns True on success, False on any error.
function DownloadEasyComDll(const AppDir: String): Boolean;
var
  PS1, Script: String;
  Lines: TArrayOfString;
  ResultCode: Integer;
begin
  Result := False;
  PS1               := ExpandConstant('{tmp}\geteasydll.ps1');
  EasyDllResultFile := ExpandConstant('{tmp}\easydll_result.txt');
  DeleteFile(EasyDllResultFile);

  Script :=
    'param([string]$AppDir, [string]$ResultFile)' + #13#10 +
    '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12' + #13#10 +
    '$base    = ' + #39 + 'https://es-assets.eaton.com/AUTOMATION/DOWNLOAD/DOWNLOADCENTER/EASY/LIB' + #39 + #13#10 +
    '$bestVer = 250' + #13#10 +
    '$bestUrl = "$base/EASY_COM_V250.zip"' + #13#10 +
    'for ($v = 251; $v -le 260; $v++) {' + #13#10 +
    '    try {' + #13#10 +
    '        $req = [Net.HttpWebRequest]::Create("$base/EASY_COM_V$v.zip")' + #13#10 +
    '        $req.Method = ' + #39 + 'HEAD' + #39 + '; $req.Timeout = 3000' + #13#10 +
    '        ($req.GetResponse()).Close()' + #13#10 +
    '        $bestVer = $v; $bestUrl = "$base/EASY_COM_V$v.zip"' + #13#10 +
    '    } catch {}' + #13#10 +
    '}' + #13#10 +
    '$zipPath    = [IO.Path]::Combine($env:TEMP, "EASY_COM_V$bestVer.zip")' + #13#10 +
    '$extractDir = [IO.Path]::Combine($AppDir, "EASY_COM_V$bestVer")' + #13#10 +
    'try {' + #13#10 +
    '    (New-Object Net.WebClient).DownloadFile($bestUrl, $zipPath)' + #13#10 +
    '    if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }' + #13#10 +
    '    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force' + #13#10 +
    '    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue' + #13#10 +
    '    $dll = Get-ChildItem -Path $extractDir -Filter ' + #39 + 'EASY_COM.dll' + #39 + ' -Recurse |' + #13#10 +
    '           Select-Object -First 1' + #13#10 +
    '    if ($dll) {' + #13#10 +
    '        $rel = $dll.FullName.Substring($AppDir.TrimEnd(' + #39 + '\' + #39 + ').Length)' +
                   '.TrimStart(' + #39 + '\' + #39 + ') -replace ' + #39 + '\\' + #39 + ', ' + #39 + '/' + #39 + #13#10 +
    '        [IO.File]::WriteAllLines($ResultFile,' + #13#10 +
    '            @([string]$bestVer, $dll.FullName, $rel, $extractDir))' + #13#10 +
    '    } else {' + #13#10 +
    '        Set-Content $ResultFile ' + #39 + 'ERROR:DLL_NOT_FOUND' + #39 + ' -Encoding ASCII' + #13#10 +
    '    }' + #13#10 +
    '} catch {' + #13#10 +
    '    Set-Content $ResultFile "ERROR:$($_.Exception.Message)" -Encoding ASCII' + #13#10 +
    '}';

  SaveStringToFile(PS1, Script, False);

  WizardForm.StatusLabel.Caption := 'Suche aktuelle EASY_COM.dll auf Eaton-Server ...';

  Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -File "' + PS1 + '"' +
    ' -AppDir "' + AppDir + '" -ResultFile "' + EasyDllResultFile + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  DeleteFile(PS1);

  if not LoadStringsFromFile(EasyDllResultFile, Lines) then Exit;
  if (GetArrayLength(Lines) = 0) or (Pos('ERROR:', Lines[0]) = 1) then Exit;

  Result := True;
end;

// Searches ExtractDir (and one subdirectory level) for readme*.txt / liesmich*.txt.
// If found, asks the user and opens the files with the system viewer.
procedure OfferReadme(const ExtractDir: String);
var
  FindRec, Sub: TFindRec;
  Paths: TArrayOfString;
  Count, ErrCode, I: Integer;
  Name, List: String;
begin
  Count := 0;
  SetArrayLength(Paths, 20);

  // Top-level .txt files
  if FindFirst(ExtractDir + '\*.txt', FindRec) then
  begin
    try
      repeat
        Name := LowerCase(FindRec.Name);
        if ((Pos('readme', Name) = 1) or (Pos('liesmich', Name) = 1)) and (Count < 20) then
        begin
          Paths[Count] := ExtractDir + '\' + FindRec.Name;
          Count := Count + 1;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
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
          if FindFirst(ExtractDir + '\' + FindRec.Name + '\*.txt', Sub) then
          begin
            try
              repeat
                Name := LowerCase(Sub.Name);
                if ((Pos('readme', Name) = 1) or (Pos('liesmich', Name) = 1)) and (Count < 20) then
                begin
                  Paths[Count] := ExtractDir + '\' + FindRec.Name + '\' + Sub.Name;
                  Count := Count + 1;
                end;
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

  if Count = 0 then Exit;

  List := '';
  for I := 0 to Count - 1 do
  begin
    if List <> '' then List := List + ', ';
    List := List + ExtractFileName(Paths[I]);
  end;

  if MsgBox('Die EASY_COM-Bibliothek wurde erfolgreich installiert.' + #13#10 +
            'Gefundene Dokumentation: ' + List + #13#10#13#10 +
            'Jetzt anzeigen?',
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

  if MsgBox('EASY_COM.dll manuell von der Festplatte auswählen?' + #13#10 +
            'Die Datei wird in das Installationsverzeichnis kopiert.',
            mbConfirmation, MB_YESNO) <> IDYES then Exit;

  PS1        := ExpandConstant('{tmp}\browsedll.ps1');
  ResultFile := ExpandConstant('{tmp}\browsedll_result.txt');
  DeleteFile(ResultFile);

  Script :=
    'Add-Type -AssemblyName System.Windows.Forms' + #13#10 +
    '$dlg = New-Object System.Windows.Forms.OpenFileDialog' + #13#10 +
    '$dlg.Title  = ' + #39 + 'EASY_COM.dll auswählen' + #39 + #13#10 +
    '$dlg.Filter = ' + #39 + 'EASY_COM.dll|EASY_COM.dll|DLL-Dateien (*.dll)|*.dll' + #39 + #13#10 +
    '$dlg.FileName = ' + #39 + 'EASY_COM.dll' + #39 + #13#10 +
    'if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {' + #13#10 +
    '    Set-Content "' + ResultFile + '" $dlg.FileName -Encoding ASCII' + #13#10 +
    '}';

  SaveStringToFile(PS1, Script, False);

  // -STA required for WinForms dialogs; SW_HIDE suppresses the console window
  Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -STA -File "' + PS1 + '"',
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
    MsgBox('Fehler beim Kopieren der Datei:' + #13#10 + SrcPath, mbError, MB_OK);
    Exit;
  end;

  IniFile := AppDir + '\easycom.ini';
  SetIniString('global', 'dll_path', 'EASY_COM.dll', IniFile);

  Result := True;
end;

// ────────────────────────────────────────────────────────────────────────────

procedure InitializeWizard;
begin
  PortPage := CreateInputQueryPage(wpSelectDir,
    'Port and Interface Configuration',
    'Configure the ports and the COM interface.',
    '');
  PortPage.Add('HTTP port (web console):', False);
  PortPage.Add('Telnet port:', False);
  PortPage.Add('COM port (1=COM1, 2=COM2 ...):', False);
  PortPage.Add('Baud rate:', False);
  PortPage.Values[0] := '8083';
  PortPage.Values[1] := '23';
  PortPage.Values[2] := '1';
  PortPage.Values[3] := '9600';

  AuthPage := CreateInputQueryPage(PortPage.ID,
    'Access Protection (optional)',
    'HTTP Basic Authentication for the web console.',
    'Leave blank to disable authentication.');
  AuthPage.Add('Username:', False);
  AuthPage.Add('Password:', True);
  AuthPage.Values[0] := 'admin';
  AuthPage.Values[1] := '';
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
      MsgBox('Please enter a valid HTTP port (1-65535).', mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Telnet < 1) or (Telnet > 65535) then
    begin
      MsgBox('Please enter a valid Telnet port (1-65535).', mbError, MB_OK);
      Result := False; Exit;
    end;
    if Http = Telnet then
    begin
      MsgBox('HTTP port and Telnet port must be different.', mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Com < 1) or (Com > 32) then
    begin
      MsgBox('Please enter a valid COM port number (1-32).', mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Baud < 300) or (Baud > 115200) then
    begin
      MsgBox('Please enter a valid baud rate (300-115200).', mbError, MB_OK);
      Result := False; Exit;
    end;
  end;

  if CurPageID = AuthPage.ID then
  begin
    if (AuthPage.Values[0] <> '') and (AuthPage.Values[1] = '') then
    begin
      if MsgBox('No password specified.' + #13#10 +
                'Without a password the web console is unprotected.' + #13#10 +
                'Continue anyway?',
                mbConfirmation, MB_YESNO) = IDNO then
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
    HttpPort   := PortPage.Values[0];
    TelnetPort := PortPage.Values[1];
    ComPort    := PortPage.Values[2];
    BaudRate   := PortPage.Values[3];
    AuthUser   := AuthPage.Values[0];
    AuthPass   := AuthPage.Values[1];

    // Write settings to easycom.ini
    IniFile := ExpandConstant('{app}\easycom.ini');
    SetIniString('instance', 'http_port',   HttpPort,   IniFile);
    SetIniString('instance', 'telnet_port', TelnetPort, IniFile);
    SetIniString('instance', 'com_port',    ComPort,    IniFile);
    SetIniString('instance', 'baud_rate',   BaudRate,   IniFile);

    if (AuthUser <> '') and (AuthPass <> '') then
    begin
      SetIniString('instance', 'basic_auth', 'true',                      IniFile);
      SetIniString('instance', 'auth_user',  AuthUser,                    IniFile);
      SetIniString('instance', 'auth_pass',  HashPasswordSha256(AuthPass), IniFile);
    end
    else
      SetIniString('instance', 'basic_auth', 'false', IniFile);

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
    if DownloadEasyComDll(ExpandConstant('{app}')) then
    begin
      if LoadStringsFromFile(EasyDllResultFile, DllLines) and
         (GetArrayLength(DllLines) >= 4) then
      begin
        // DllLines[2] = relative path with forward slashes (e.g. EASY_COM_V250/EASY_COM.dll)
        SetIniString('global', 'dll_path', DllLines[2], IniFile);
        OfferReadme(DllLines[3]);
      end;
      DeleteFile(EasyDllResultFile);
    end
    else
    begin
      // Download fehlgeschlagen — manuelle Auswahl anbieten
      MsgBox('EASY_COM.dll konnte nicht automatisch heruntergeladen werden.' + #13#10 +
             '(Kein Internetzugang oder Eaton-Server nicht erreichbar.)',
             mbInformation, MB_OK);
      if not BrowseForEasyComDll(ExpandConstant('{app}')) then
        MsgBox('EASY_COM.dll wurde nicht installiert.' + #13#10#13#10 +
               'Bitte laden Sie die Bibliothek manuell herunter:' + #13#10 +
               'https://es-assets.eaton.com/AUTOMATION/DOWNLOAD/DOWNLOADCENTER/EASY/LIB/EASY_COM_V250.zip' + #13#10#13#10 +
               'Kopieren Sie EASY_COM.dll anschließend in:' + #13#10 +
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
  end;

  if CurUninstallStep = usPostUninstall then
  begin
    // Remove Start Menu shortcut
    ShellLink := ExpandConstant('{commonprograms}\EasyComServer\EasyComServer Web Console.url');
    DeleteFile(ShellLink);
    RemoveDir(ExpandConstant('{commonprograms}\EasyComServer'));
  end;
end;
