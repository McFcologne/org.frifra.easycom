# EasyComServer — Visual Studio Project

HTTP and Telnet gateway for **EASY_COM.dll** (Moeller/Eaton EASY PLCs).  
Runs as a **Windows Service** on Windows Server 2025/2022/2019.

---

## Quick Start in Visual Studio

### 1. Prerequisites

| | |
|---|---|
| Visual Studio | 2022 (Community edition is sufficient) |
| Workload | **.NET Desktop Development** or **ASP.NET and web development** |
| .NET SDK | 8.0 |
| Target platform | **x86** — EASY_COM.dll is 32-bit |

### 2. Open the project

```
Double-click  EasyComServer.sln
```

### 3. Add EASY_COM.dll

Copy the 32-bit DLL into the project folder:

```
EasyComServer\
  EasyComServer\
    EASY_COM.dll   ← here
    EasyComServer.csproj
    ...
```

Then in Visual Studio: **Right-click the project → Add → Existing Item → EASY_COM.dll**  
File properties: *Copy to Output Directory* → **Copy if newer**

### 4. Check the platform

Menu: **Build → Configuration Manager**  
Make sure **x86** is set as the platform (not AnyCPU!).

### 5. Debugging (console mode)

`F5` starts the application with the **"Console (Debug)"** profile,  
i.e. `EasyComServer.exe --console`. Server output appears in the  
console window — no service registration required.

The launch profile can be changed in `Properties\launchSettings.json`.

### 6. Adjust configuration

Edit `easycom.ini` in the project folder — it is automatically copied  
to the output directory on each build.

```ini
[instance]
com_port    = 1       ; ← COM port of the EASY device
baud_rate   = 9600    ; ← baud rate
http_port   = 8083
telnet_port = 8023
basic_auth  = false   ; true = enable password protection
```

---

## Build & Deployment

### Create a release build

```
Menu: Build → Publish EasyComServer
```

or from the command line:

```powershell
dotnet publish -c Release -r win-x86 --self-contained false -o .\publish
```

> **Note:** `EASY_COM.dll` must be copied manually into the `publish\` folder
> if it was not already in the project folder.

### Install as a Windows Service

```powershell
# Run as Administrator
.\publish\install-service.ps1 -Action install
```

The script:
- Registers the service with auto-start
- Configures automatic restart on crash
- Opens firewall ports (HTTP + Telnet)
- Registers HTTP URL ACLs

### Service management

```powershell
.\install-service.ps1 -Action start     # Start
.\install-service.ps1 -Action stop      # Stop
.\install-service.ps1 -Action status    # Show status
.\install-service.ps1 -Action uninstall
```

---

## Web Console

After starting, open in your browser:

```
http://localhost:8083/
```

→ Redirects to `http://localhost:8083/index.html`  
→ Interactive terminal console with command history, tab completion and quick buttons

---

## REST API

The server exposes a versioned REST API at `/api/v1/` with JSON responses.

```
http://SERVER:PORT/api/v1/system
http://SERVER:PORT/api/v1/connection
http://SERVER:PORT/api/v1/device/{netId}/clock
http://SERVER:PORT/api/v1/device/{netId}/objects/{obj}/{index}
```

All responses follow a uniform envelope:

```json
{ "ok": true,  "result": "..." }
{ "ok": false, "error":  "..." }
```

The full API specification (OpenAPI 3.0) is served at:

```
http://SERVER:PORT/openapi.yaml
```

An interactive Swagger UI test page is available at:

```
http://localhost:8083/swagger.html
```

---

## HTTP API (legacy)

```
http://SERVER:PORT/easy.cmd?COMMAND
```

Examples:
```
http://localhost:8083/easy.cmd?help
http://localhost:8083/easy.cmd?show%20server
http://localhost:8083/easy.cmd?read_clock%200
http://localhost:8083/easy.cmd?read_object_value%200%201%200%204
```

Commands can be abbreviated (`sh ser` = `show server`).

---

## Project Structure

```
EasyComServer.sln
EasyComServer\
  EasyComServer.csproj      Project file (x86, .NET 8, Windows Service)
  EasyComServer.cs          ServiceBase + HTTP/Telnet listeners
  EasyComWrapper.cs         P/Invoke wrapper for EASY_COM.dll + auto-connect
  CommandProcessor.cs       Command parser, dispatcher, abbreviation logic
  ServerConfig.cs           INI file reader/writer
  Logger.cs                 Thread-safe file and console logger
  easycom.ini               Configuration (copied to build directory automatically)
  install-service.ps1       PowerShell installation script
  EASY_COM.dll              ← insert here (32-bit, not in git)
  wwwroot\
    index.html              Web console
  Properties\
    launchSettings.json     VS debug profiles
```

---

## Common Errors

| Error | Cause | Solution |
|---|---|---|
| `BadImageFormatException` | Project built as x64/AnyCPU | Set platform to **x86** |
| `DllNotFoundException` | EASY_COM.dll missing from output directory | Copy DLL to project folder, set `CopyToOutputDirectory = PreserveNewest` |
| `HttpListenerException: Access denied` | No HTTP URL ACL registered | Run `install-service.ps1` as Admin, or run `netsh http add urlacl` manually |
| `Access is denied` (COM port) | Port held by another process | Stop the other process; reduce idle_timeout |
---

## Versions

| Version | Date | Description |
|---|---|---|
| 0.0.1 | 2009-07-21 | Initial prototype — Windows Service implementation for Windows Server 2003 |
| 1.0.0 | 2009-08-05 | First stable release of the Delphi-based server application |
| 2.0.0 | 2026-04-07 | Full rewrite in C# (.NET 8) — migrated from legacy Delphi codebase to Visual Studio; added HTTP/Telnet gateway, multi-instance support, web console and Basic Auth |
| 2.1.0 | 2026-04-08 | Added RESTful API (`/api/v1`), OpenAPI 3.0 specification (`openapi.yaml`), interactive Swagger UI test page (`/swagger.html`) |

---

## Background

This project grew out of my own use of Moeller EASY PLCs in a home automation and building control setup. The devices are connected via RS232-to-Ethernet adapters, which expose virtual COM ports on a Windows Server — making the PLCs accessible over the network just like a locally attached serial device.

I use this setup in two ways: with **EasySoft** for programming and extending the PLC logic, and with this HTTP gateway for runtime monitoring and control. Having a simple HTTP API means I can trigger outputs, read inputs and check device state from any script, browser or home automation system without needing EasySoft running.

The original gateway was a Delphi application written in 2009. After running reliably for over 15 years, it was rewritten in C# (.NET 8) to make it easier to maintain, extend and deploy on modern Windows Server versions.
