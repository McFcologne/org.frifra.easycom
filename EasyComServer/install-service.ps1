# install-service.ps1
# Run as Administrator on Windows Server 2025
# Usage:
#   .\install-service.ps1 -Action install
#   .\install-service.ps1 -Action uninstall

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("install","uninstall","start","stop","status")]
    [string]$Action,

    [string]$ServiceName = "EasyComServer",
    [string]$DisplayName = "Moeller EASY COM Server",
    [string]$Description = "HTTP and Telnet gateway for EASY_COM.dll. Enables remote access to Moeller EASY devices.",
    [string]$ExePath = "$PSScriptRoot\EasyComServer.exe",
    [string]$User = "LocalSystem"
)

$ErrorActionPreference = "Stop"

function Require-Admin {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
              ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "This script must be run as Administrator."
        exit 1
    }
}

switch ($Action) {
    "install" {
        Require-Admin
        if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
            Write-Host "Service '$ServiceName' already exists. Stopping and removing first..."
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            sc.exe delete $ServiceName | Out-Null
            Start-Sleep 2
        }

        Write-Host "Installing service '$ServiceName'..."
        New-Service `
            -Name $ServiceName `
            -DisplayName $DisplayName `
            -Description $Description `
            -BinaryPathName "`"$ExePath`"" `
            -StartupType Automatic

        # Add failure recovery: restart after 30s on first/second failure, reboot on third
        sc.exe failure $ServiceName reset= 86400 actions= restart/30000/restart/30000/reboot/60000 | Out-Null

        # Set the service description (New-Service -Description may not persist on older PS)
        sc.exe description $ServiceName $Description | Out-Null

        Write-Host "Opening firewall for HTTP and Telnet ports..."
        # Read ports from ini if possible
        $iniPath = Join-Path (Split-Path $ExePath -Parent) "easycom.ini"
        $httpPorts = @(8083)
        $telnetPorts = @(8023)
        if (Test-Path $iniPath) {
            $httpPorts  = Select-String -Path $iniPath -Pattern "^\s*http_port\s*=\s*(\d+)" |
                          ForEach-Object { $_.Matches.Groups[1].Value } | Select-Object -Unique
            $telnetPorts = Select-String -Path $iniPath -Pattern "^\s*telnet_port\s*=\s*(\d+)" |
                           ForEach-Object { $_.Matches.Groups[1].Value } | Select-Object -Unique
        }

        foreach ($p in $httpPorts) {
            netsh advfirewall firewall add rule name="EasyComServer HTTP $p" `
                dir=in action=allow protocol=TCP localport=$p 2>&1 | Out-Null
        }
        foreach ($p in $telnetPorts) {
            netsh advfirewall firewall add rule name="EasyComServer Telnet $p" `
                dir=in action=allow protocol=TCP localport=$p 2>&1 | Out-Null
        }

        # Register HTTP URL ACL so the service can listen without admin at runtime
        foreach ($p in $httpPorts) {
            netsh http add urlacl url="http://+:$p/" user="$User" 2>&1 | Out-Null
            Write-Host "  HTTP ACL registered for port $p"
        }

        Write-Host "Starting service..."
        Start-Service -Name $ServiceName
        Write-Host "Done. Service status:"
        Get-Service -Name $ServiceName | Select-Object Name, Status, StartType
    }

    "uninstall" {
        Require-Admin
        Write-Host "Stopping '$ServiceName'..."
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Write-Host "Removing '$ServiceName'..."
        sc.exe delete $ServiceName | Out-Null
        Write-Host "Removing firewall rules..."
        netsh advfirewall firewall delete rule name="EasyComServer HTTP 8083" 2>&1 | Out-Null
        netsh advfirewall firewall delete rule name="EasyComServer Telnet 8023" 2>&1 | Out-Null
        Write-Host "Removing URL ACLs..."
        netsh http delete urlacl url="http://+:8083/" 2>&1 | Out-Null
        Write-Host "Uninstalled."
    }

    "start"  { Start-Service  -Name $ServiceName; Get-Service -Name $ServiceName }
    "stop"   { Stop-Service   -Name $ServiceName; Get-Service -Name $ServiceName }
    "status" { Get-Service    -Name $ServiceName | Format-List * }
}
