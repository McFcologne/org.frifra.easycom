[Setup]
AppName=EasyComServer
AppVersion=2.0.0
AppPublisher=Michael Fritzsche
DefaultDirName={autopf}\EasyComServer
DefaultGroupName=EasyComServer
OutputDir=.
OutputBaseFilename=EasyComServer_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Files]
Source: "..\bin\Release\net8.0-windows\publish\*"; DestDir: "{app}"; Flags: recursesubdirs

[Run]
Filename: "sc"; Parameters: "create EasyComServer binPath=""{app}\EasyComServer.exe"" start=auto obj=""LocalSystem"" DisplayName=""Moeller EASY COM Server"""; Flags: runhidden
Filename: "sc"; Parameters: "description EasyComServer ""HTTP und Telnet Gateway fuer EASY_COM.dll"""; Flags: runhidden
Filename: "sc"; Parameters: "start EasyComServer"; Flags: runhidden

[UninstallRun]
Filename: "sc"; Parameters: "stop EasyComServer"; Flags: runhidden; RunOnceId: "StopService"
Filename: "cmd"; Parameters: "/c timeout /t 3 /nobreak"; Flags: runhidden; RunOnceId: "Wait"
Filename: "sc"; Parameters: "delete EasyComServer"; Flags: runhidden; RunOnceId: "DeleteService"

[Icons]
Name: "{group}\Deinstallieren"; Filename: "{uninstallexe}"

[Code]
var
  PortPage: TInputQueryWizardPage;
  AuthPage: TInputQueryWizardPage;

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
  WizardForm.StatusLabel.Caption := '.NET 8 Runtime wird heruntergeladen...';

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
    MsgBox('Download fehlgeschlagen.' + #13#10 +
           'Bitte .NET 8 Runtime (x86) manuell installieren:' + #13#10 +
           'https://aka.ms/dotnet/8.0/dotnet-runtime-win-x86.exe',
           mbError, MB_OK);
    Exit;
  end;

  WizardForm.StatusLabel.Caption := '.NET 8 Runtime wird installiert...';
  if Exec(TempFile, '/quiet /norestart', '', SW_HIDE,
          ewWaitUntilTerminated, ResultCode) then
  begin
    DeleteFile(TempFile);
    if (ResultCode = 0) or (ResultCode = 1638) then
      Result := True
    else
      MsgBox('.NET Installation fehlgeschlagen (Code: ' +
             IntToStr(ResultCode) + ')', mbError, MB_OK);
  end
  else
    MsgBox('.NET Installer konnte nicht gestartet werden.', mbError, MB_OK);
end;

procedure InitializeWizard;
begin
  PortPage := CreateInputQueryPage(wpSelectDir,
    'Port- und Schnittstellenkonfiguration',
    'Legen Sie die Ports und den COM-Anschluss fest.',
    '');
  PortPage.Add('HTTP-Port (Web-Konsole):', False);
  PortPage.Add('Telnet-Port:', False);
  PortPage.Add('COM-Port (1=COM1, 2=COM2 ...):', False);
  PortPage.Add('Baudrate:', False);
  PortPage.Values[0] := '8083';
  PortPage.Values[1] := '8023';
  PortPage.Values[2] := '1';
  PortPage.Values[3] := '9600';

  AuthPage := CreateInputQueryPage(PortPage.ID,
    'Zugangssicherung (optional)',
    'HTTP Basic Authentication fuer die Web-Konsole.',
    'Leer lassen um Authentifizierung zu deaktivieren.');
  AuthPage.Add('Benutzername:', False);
  AuthPage.Add('Passwort:', True);
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
      MsgBox('Bitte einen gueltigen HTTP-Port eingeben (1-65535).', mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Telnet < 1) or (Telnet > 65535) then
    begin
      MsgBox('Bitte einen gueltigen Telnet-Port eingeben (1-65535).', mbError, MB_OK);
      Result := False; Exit;
    end;
    if Http = Telnet then
    begin
      MsgBox('HTTP-Port und Telnet-Port duerfen nicht identisch sein.', mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Com < 1) or (Com > 32) then
    begin
      MsgBox('Bitte einen gueltigen COM-Port eingeben (1-32).', mbError, MB_OK);
      Result := False; Exit;
    end;
    if (Baud < 300) or (Baud > 115200) then
    begin
      MsgBox('Bitte eine gueltige Baudrate eingeben (300-115200).', mbError, MB_OK);
      Result := False; Exit;
    end;
  end;

  if CurPageID = AuthPage.ID then
  begin
    if (AuthPage.Values[0] <> '') and (AuthPage.Values[1] = '') then
    begin
      if MsgBox('Kein Passwort angegeben.' + #13#10 +
                'Ohne Passwort ist die Web-Konsole ungeschuetzt.' + #13#10 +
                'Trotzdem fortfahren?',
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
      if MsgBox('.NET 8 Runtime (x86) ist nicht installiert.' + #13#10 +
                'Soll sie jetzt heruntergeladen und installiert werden? (~26 MB)',
                mbConfirmation, MB_YESNO) = IDYES then
      begin
        if not DownloadAndInstallDotNet then
        begin
          Result := False; Exit;
        end;
      end
      else
      begin
        MsgBox('EasyComServer benoetigt .NET 8 Runtime (x86).' + #13#10 +
               'Die Installation wird abgebrochen.', mbError, MB_OK);
        Result := False; Exit;
      end;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  IniFile, HtmlFile, HtmlContent: String;
  HttpPort, TelnetPort, ComPort, BaudRate: String;
  AuthUser, AuthPass: String;
  ShellDir, ShellLink: String;
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

    // easycom.ini schreiben
    IniFile := ExpandConstant('{app}\easycom.ini');
    SetIniString('instance', 'http_port',   HttpPort,   IniFile);
    SetIniString('instance', 'telnet_port', TelnetPort, IniFile);
    SetIniString('instance', 'com_port',    ComPort,    IniFile);
    SetIniString('instance', 'baud_rate',   BaudRate,   IniFile);

    if (AuthUser <> '') and (AuthPass <> '') then
    begin
      SetIniString('instance', 'basic_auth', 'true',   IniFile);
      SetIniString('instance', 'auth_user',  AuthUser, IniFile);
      SetIniString('instance', 'auth_pass',  AuthPass, IniFile);
    end
    else
      SetIniString('instance', 'basic_auth', 'false', IniFile);

     // index.html patchen: Port via PS-Skript
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

    // URL ACL registrieren
    Exec('netsh',
      'http add urlacl url="http://+:' + HttpPort + '/" user="Everyone"',
      '', SW_HIDE, ewWaitUntilTerminated, Dummy);

    // Startmenue-Shortcut anlegen
    ShellDir  := ExpandConstant('{commonprograms}\EasyComServer');
    ShellLink := ShellDir + '\EasyComServer Web Console.url';
    ForceDirectories(ShellDir);
    SaveStringToFile(ShellLink,
      '[InternetShortcut]' + #13#10 +
      'URL=http://localhost:' + HttpPort + '/' + #13#10,
      False);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ShellLink, HttpPort, IniFile: String;
  Dummy: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    IniFile  := ExpandConstant('{app}\easycom.ini');
    HttpPort := GetIniString('instance', 'http_port', '8083', IniFile);
    Exec('netsh',
      'http delete urlacl url="http://+:' + HttpPort + '/"',
      '', SW_HIDE, ewWaitUntilTerminated, Dummy);
  end;

  if CurUninstallStep = usPostUninstall then
  begin
    ShellLink := ExpandConstant('{commonprograms}\EasyComServer\EasyComServer Web Console.url');
    DeleteFile(ShellLink);
    RemoveDir(ExpandConstant('{commonprograms}\EasyComServer'));
  end;
end;
