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

## HTTP API

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
