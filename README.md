# EasyComServer — Visual Studio Projekt

HTTP- und Telnet-Gateway für die **EASY_COM.dll** (Moeller/Eaton EASY-Steuerungen).  
Läuft als **Windows-Service** unter Windows Server 2025/2022/2019.

---

## Schnellstart in Visual Studio

### 1. Voraussetzungen

| | |
|---|---|
| Visual Studio | 2022 (Community reicht) |
| Workload | **.NET Desktop Development** oder **ASP.NET and web development** |
| .NET SDK | 8.0 |
| Zielplattform | **x86** — EASY_COM.dll ist 32-Bit |

### 2. Projekt öffnen

```
Doppelklick auf  EasyComServer.sln
```

### 3. EASY_COM.dll einfügen

Die 32-Bit-DLL in den Projektordner kopieren:

```
EasyComServer\
  EasyComServer\
    EASY_COM.dll   ← hier
    EasyComServer.csproj
    ...
```

Danach in Visual Studio: **Rechtsklick auf das Projekt → Hinzufügen → Vorhandenes Element → EASY_COM.dll**  
Eigenschaften der Datei: *In Ausgabeverzeichnis kopieren* → **Wenn neuer**

### 4. Plattform prüfen

Menü: **Build → Configuration Manager**  
Sicherstellen dass **x86** als Plattform gesetzt ist (nicht AnyCPU!).

### 5. Debuggen (Konsolenmodus)

`F5` startet die Anwendung mit dem Profil **„Konsole (Debug)"**,  
d.h. `EasyComServer.exe --console`. Die Server-Ausgabe erscheint im  
Konsolenfenster, kein Service-Registrierung nötig.

Das Startprofil kann in `Properties\launchSettings.json` geändert werden.

### 6. Konfiguration anpassen

`easycom.ini` im Projektordner bearbeiten — wird beim Build automatisch  
ins Ausgabeverzeichnis kopiert.

```ini
[instance]
com_port   = 1        ; ← COM-Port des EASY-Geräts
baud_rate  = 9600     ; ← Baudrate
http_port  = 8083
telnet_port = 8023
basic_auth = false    ; true = Passwortschutz aktivieren
```

---

## Build & Deployment

### Release-Build erstellen

```
Menü: Build → Publish EasyComServer
```

oder auf der Kommandozeile:

```powershell
dotnet publish -c Release -r win-x86 --self-contained false -o .\publish
```

> **Hinweis:** `EASY_COM.dll` muss manuell in den `publish\`-Ordner kopiert werden,
> falls sie noch nicht im Projektordner lag.

### Als Windows-Service installieren

```powershell
# Als Administrator ausführen
.\publish\install-service.ps1 -Action install
```

Das Skript:
- Registriert den Dienst mit Autostart
- Konfiguriert automatischen Neustart bei Absturz
- Öffnet Firewall-Ports (HTTP + Telnet)
- Registriert HTTP URL ACLs

### Service-Verwaltung

```powershell
.\install-service.ps1 -Action start    # Starten
.\install-service.ps1 -Action stop     # Stoppen
.\install-service.ps1 -Action status   # Status anzeigen
.\install-service.ps1 -Action uninstall
```

---

## Web-Konsole

Nach dem Start im Browser aufrufen:

```
http://localhost:8083/
```

→ Redirect auf `http://localhost:8083/index.html`  
→ Interaktive Terminal-Konsole mit Befehlsverlauf, Tab-Completion und Quick-Buttons

---

## HTTP API

```
http://SERVER:PORT/easy.cmd?BEFEHL
```

Beispiele:
```
http://localhost:8083/easy.cmd?help
http://localhost:8083/easy.cmd?show%20server
http://localhost:8083/easy.cmd?read_clock%200
http://localhost:8083/easy.cmd?read_object_value%200%201%200%204
```

Befehle können abgekürzt werden (`sh ser` = `show server`).

---

## Projektstruktur

```
EasyComServer.sln
EasyComServer\
  EasyComServer.csproj      Projektdatei (x86, .NET 8, Windows Service)
  EasyComServer.cs          ServiceBase + HTTP/Telnet-Listener
  EasyComWrapper.cs         P/Invoke-Wrapper für EASY_COM.dll + Auto-Connect
  CommandProcessor.cs       Befehls-Parser, Dispatcher, Abkürzungslogik
  ServerConfig.cs           INI-Datei lesen/schreiben
  Logger.cs                 Thread-sicherer Datei- und Konsolen-Logger
  easycom.ini               Konfiguration (wird ins Build-Verzeichnis kopiert)
  install-service.ps1       PowerShell-Installationsscript
  EASY_COM.dll              ← hier einfügen (32-Bit, nicht in git)
  wwwroot\
    index.html              Web-Konsole
  Properties\
    launchSettings.json     VS-Debugprofile
```

---

## Häufige Fehler

| Fehler | Ursache | Lösung |
|---|---|---|
| `BadImageFormatException` | Projekt als x64/AnyCPU gebaut | Platform auf **x86** setzen |
| `DllNotFoundException` | EASY_COM.dll fehlt im Ausgabeverzeichnis | DLL in Projektordner kopieren, `CopyToOutputDirectory = PreserveNewest` |
| `HttpListenerException: Access denied` | Kein HTTP URL ACL | `install-service.ps1` als Admin ausführen, oder `netsh http add urlacl` manuell |
| `Access is denied` (COM-Port) | Port wird von anderem Prozess gehalten | Anderen Prozess beenden; idle_timeout verkürzen |
