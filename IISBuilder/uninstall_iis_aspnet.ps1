#Requires -RunAsAdministrator

function Write-LogMessage {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}


function Write-LogError {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $Message" -ForegroundColor Red
}

Write-LogMessage "Starting IIS uninstallation..."

try {

    Write-LogMessage "Stopping IIS service..."
    Stop-Service -Name W3SVC -Force -ErrorAction SilentlyContinue
    Stop-Service -Name WAS -Force -ErrorAction SilentlyContinue
    

    Write-LogMessage "Uninstalling IIS features..."
    $result = Uninstall-WindowsFeature -Name Web-Server, `
        Web-Common-Http, `
        Web-Default-Doc, `
        Web-Dir-Browsing, `
        Web-Http-Errors, `
        Web-Static-Content, `
        Web-Http-Logging, `
        Web-Stat-Compression, `
        Web-Filtering, `
        Web-Mgmt-Console, `
        Web-Asp-Net45, `
        Web-ISAPI-Ext, `
        Web-ISAPI-Filter, `
        Web-Net-Ext45 -IncludeManagementTools
    
    if ($result.Success) {
        Write-LogMessage "IIS features uninstalled successfully"
    } else {
        Write-LogError "Failed to uninstall some IIS features"
    }
    
    Write-LogMessage "Cleaning up IIS directories..."
    $directories = @(
        "%SystemRoot%\System32\inetsrv",
        "%SystemRoot%\System32\LogFiles\W3SVC*",
        "%SystemRoot%\System32\config\systemprofile\AppData\Local\Temp\IIS Temporary Compressed Files"
    )
    
    foreach ($dir in $directories) {
        $expandedPath = [System.Environment]::ExpandEnvironmentVariables($dir)
        if (Test-Path $expandedPath) {
            Remove-Item -Path $expandedPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-LogMessage "Removed directory: $expandedPath"
        }
    }
    
    $webShellsPath = Read-Host "Enter the path to your web shells directory to remove it (or press Enter to skip)"
    if ($webShellsPath -and (Test-Path $webShellsPath)) {
        $confirm = Read-Host "Are you sure you want to remove $webShellsPath and all its contents? (y/n)"
        if ($confirm -eq 'y') {
            Remove-Item -Path $webShellsPath -Recurse -Force
            Write-LogMessage "Removed web shells directory: $webShellsPath"
        }
    }
    
    Write-LogMessage "Cleaning up registry..."
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\InetStp",
        "HKLM:\SOFTWARE\Microsoft\WebManagement",
        "HKLM:\SYSTEM\CurrentControlSet\Services\W3SVC",
        "HKLM:\SYSTEM\CurrentControlSet\Services\WAS"
    )
    
    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-LogMessage "Removed registry key: $path"
        }
    }
    
    Write-LogMessage "Uninstallation completed successfully!"
    
    Write-Host "`n=== Final Status ===" -ForegroundColor Green
    
    $remainingServices = @('W3SVC', 'WAS') | ForEach-Object {
        Get-Service -Name $_ -ErrorAction SilentlyContinue
    }
    
    if ($remainingServices) {
        Write-Host "Some IIS services still present (will be removed after reboot)" -ForegroundColor Yellow
    } else {
        Write-Host "All IIS services removed" -ForegroundColor Green
    }
    
    if ($result.RestartNeeded -eq 'Yes') {
        Write-Host "`nSystem restart required to complete uninstallation" -ForegroundColor Yellow
        $restart = Read-Host "Would you like to restart now? (y/n)"
        if ($restart -eq 'y') {
            Write-LogMessage "Restarting computer..."
            Restart-Computer -Force
        } else {
            Write-LogMessage "Please restart your computer to complete the uninstallation"
        }
    }
    
} catch {
    Write-LogError "Uninstallation encountered an error: $_"
    Write-LogError "Stack Trace: $($_.ScriptStackTrace)"
    
    Write-Host "`nDiagnostic Information:" -ForegroundColor Yellow
    Write-Host "- IIS Service Present: $((Get-Service W3SVC -ErrorAction SilentlyContinue) -ne $null)"
    Write-Host "- IIS Features Present: $((Get-WindowsFeature -Name Web-Server).Installed)"
    Write-Host "- Registry Keys Remaining: $(Test-Path 'HKLM:\SOFTWARE\Microsoft\InetStp')"
    
    exit 1
}