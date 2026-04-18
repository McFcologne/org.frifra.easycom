# Release Notes — EasyComServer

---

## 2.2.0 — 2026-04-19

Ab dieser Version steht ein fertiges Windows-Setup (`EasyComServer_Setup.exe`) als Download auf GitHub zur Verfügung:  
**[github.com/McFCologne/EasyComServer/releases](https://github.com/McFCologne/EasyComServer/releases)**

### Installer (setup.iss)

- **Automatischer DLL-Download:** Das Setup lädt `EASY_COM.dll` jetzt selbst von den Eaton-Servern herunter. Es prüft zunächst, ob eine neuere Version als V250 verfügbar ist (HEAD-Anfragen V251–V260, 3 s Timeout je Version), und nimmt die aktuellste gefundene. Schlägt der Abruf fehl, wird auf die statische V250-URL zurückgefallen.
- **Versionierte Ablage:** Das heruntergeladene ZIP wird vollständig in einen eigenen Unterordner `EASY_COM_V{n}` im Programmverzeichnis entpackt. `dll_path` in `easycom.ini` wird automatisch auf den relativen Pfad der enthaltenen DLL gesetzt (z. B. `EASY_COM_V250/EASY_COM.dll`).
- **Readme-Angebot:** Nach dem Entpacken sucht das Setup nach `readme*.txt` und `liesmich*.txt` im extrahierten Ordner und bietet an, diese mit dem Standard-Viewer zu öffnen.
- **Manueller Fallback:** Schlägt der Download fehl (kein Internetzugang, Server nicht erreichbar), öffnet das Setup einen Windows-Dateiauswahl-Dialog. Die gewählte DLL wird nach `{app}\EASY_COM.dll` kopiert und als `dll_path` eingetragen.
- **EASY_COM.dll aus Paket ausgeschlossen:** Die DLL darf nicht weiterverteilt werden und ist daher nicht mehr im Installer-Paket enthalten. Sie wird ausschließlich zur Laufzeit bezogen. Für lokale Debug-Builds bleibt der Build-Output-Eintrag erhalten.

### Laufzeit (EasyComServer.cs)

- **Konfigurierter DLL-Pfad wird tatsächlich genutzt:** Alle `[DllImport("EASY_COM.dll")]`-Aufrufe werden jetzt über `NativeLibrary.SetDllImportResolver` auf den in `easycom.ini` konfigurierten absoluten Pfad umgeleitet. Bisher wurde `dll_path` nur für `GetVersion()` ausgewertet; Windows suchte die DLL nach seiner eigenen Reihenfolge und ignorierte Unterverzeichnisse.
- **Klare Fehlermeldung beim Start:** Ist die konfigurierte DLL nicht vorhanden, wirft die Anwendung sofort eine `FileNotFoundException` mit Pfadangabe, statt mit einer kryptischen P/Invoke-Exception abzustürzen.

### Build (EasyComServer.csproj)

- `EASY_COM.dll`-Eintrag mit `Condition="Exists('EASY_COM.dll')"` gesichert: Der Build schlägt nicht fehl, wenn die Datei lokal nicht vorhanden ist (z. B. frischer Clone, CI-Pipeline).
- Doppelter `easycom.ini`-Eintrag entfernt.

---

## 2.1.1 — 2026-04-09

- SHA-256-Passwort-Hashing für HTTP Basic Auth (`sha256:<hex>` in `auth_pass`).
- Hilfsprogramm `EasyComServer.exe --hash-password <passwort>` zur Hash-Generierung.
- MIT-Lizenz hinzugefügt.

---

## 2.1.0 — 2026-04-08

- REST-API unter `/api/v1/` mit einheitlichem JSON-Envelope (`ok`, `result`/`error`).
- OpenAPI 3.0-Spezifikation (`openapi.yaml`), über HTTP ausgeliefert.
- Interaktive Swagger-UI-Testseite (`/swagger.html`).

---

## 2.0.0 — 2026-04-07

Vollständige Neuentwicklung in C# (.NET 8); Migration von der Delphi-Codebasis (2009).

- HTTP- und Telnet-Gateway für `EASY_COM.dll`.
- Multi-Instanz-Unterstützung: mehrere `[instance]`-Blöcke in `easycom.ini` für unabhängige COM-Ports.
- Web-Konsole (`index.html`) mit Befehlshistorie, Tab-Vervollständigung und Schnellschaltflächen.
- HTTP Basic Auth (optional, pro Instanz).
- Automatisches Öffnen/Schließen des COM-Ports (`com_idle_timeout`).
- Pulse-Kommando (`WRITE_OBJECT_VALUE … v1|v2|ms`) für zeitgesteuerte Wertpaare.
- Windows-Service-Installer (`install-service.ps1`) mit Firewall- und URL-ACL-Konfiguration.
- Inno-Setup-Skript für GUI-Installation mit Port-/Auth-Konfigurationsseiten.

---

## 1.0.0 — 2009-08-05

Erste stabile Version der Delphi-basierten Serveranwendung.

---

## 0.0.1 — 2009-07-21

Erster Prototyp — Windows-Service-Implementierung für Windows Server 2003.
