# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

```powershell
# Debug build
dotnet build EasyComServer/EasyComServer.csproj

# Release publish (self-contained=false, x86)
dotnet publish EasyComServer/EasyComServer.csproj -c Release -r win-x86 --self-contained false -o ./publish

# Run in console mode (no service registration needed)
EasyComServer.exe --console
```

**Platform:** Must be **x86** — `EASY_COM.dll` is a 32-bit Delphi DLL. AnyCPU/x64 builds will fail at runtime.

**No test suite.** Manual testing via:
- Swagger UI: `http://localhost:8083/swagger.html`
- Legacy endpoint: `http://localhost:8083/easy.cmd?help`
- Telnet: `telnet localhost 8023`

`easycom.ini` is auto-copied to the output directory on each build; edit it there or in the project root.

## Architecture

EasyComServer is a Windows Service gateway that wraps `EASY_COM.dll` (a 32-bit Delphi COM-port library for Moeller/Eaton EASY PLCs) and exposes device operations over HTTP and Telnet.

### Component Map

```
HTTP (port 8083)  ──►  EasyComService.HandleHttp()
                           │
                           ├─ /api/v1/*  ──►  RestApiHandler  ──►  CommandProcessor
                           └─ /easy.cmd? ──────────────────────►  CommandProcessor
                                                                        │
Telnet (port 8023) ─────────────────────────────────────────────►  CommandProcessor
                                                                        │
                                                                   EasyComWrapper
                                                                        │
                                                                  P/Invoke → EASY_COM.dll
                                                                        │
                                                               COM port / Ethernet → PLC
```

**`EasyComWrapper`** is the only component that touches `EASY_COM.dll`. It serializes every call through a single `_lock` — the Delphi DLL is not thread-safe and concurrent calls crash it.

**`CommandProcessor`** parses plain-text command strings (from both HTTP and Telnet) and delegates to `EasyComWrapper`. It supports prefix abbreviation: `SH SER` resolves to `SHOW SERVER`. The REST API is a thin JSON envelope — `RestApiHandler` unmarshals the request and delegates to `CommandProcessor.Execute()` with a command string.

**Multi-instance:** Each `[instance]` block in `easycom.ini` gets its own `EasyComWrapper` and its own HTTP/Telnet listener pair, allowing independent COM ports (e.g., two PLCs on COM1 and COM2 simultaneously).

### Pulse Commands

`WRITE_OBJECT_VALUE netId obj index length v1|v2|ms` writes `v1`, returns immediately (`OK PULSE ... async`), then writes `v2` after `ms` milliseconds via a background `Task`. Other requests may proceed during the wait; serial exclusivity is maintained by `_lock` inside `EasyComWrapper`. A `KeepAlive()` loop (every 5 s) prevents the idle-close timer from firing during the wait.

### COM Port Lifecycle

The COM port auto-opens before any command (`EnsureComConnected`) and auto-closes after `com_idle_timeout` seconds of inactivity (default 300 s, checked every 10 s). The next command after a close re-opens transparently.

### Key Files

| File | Purpose |
|------|---------|
| `EasyComServer.cs` | Service entry point; HTTP/Telnet listeners; per-instance task startup |
| `EasyComWrapper.cs` | P/Invoke wrapper; connection lifecycle; `_lock` serialization |
| `CommandProcessor.cs` | Command parsing, dispatch, pulse logic |
| `RestApiHandler.cs` | `/api/v1/*` routing and JSON marshaling |
| `ServerConfig.cs` | `easycom.ini` parser; `InstanceConfig` model |
| `wwwroot/openapi.yaml` | OpenAPI 3.0 spec (authoritative API reference) |

### P/Invoke Conventions

All `EASY_COM.dll` functions use `int` parameters (not `byte`/`ushort`) due to the Delphi 32-bit calling convention. Return code `0` = success; non-zero codes are looked up via `GetLastSysError` / `EasyErrorString`.
