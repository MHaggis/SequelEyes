#Requires -RunAsAdministrator

function Write-LogMessage {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Write-LogError {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $Message" -ForegroundColor Red
}

function Get-AllIPAddresses {
    $ips = Get-NetIPAddress -AddressFamily IPv4 | 
           Where-Object { $_.InterfaceAlias -notmatch 'Loopback' } |
           Select-Object IPAddress
    return $ips
}

function Test-IISInstallation {
    param(
        [int]$RetryCount = 3,
        [int]$RetryDelay = 5
    )
    
    for ($i = 1; $i -le $RetryCount; $i++) {
        Write-LogMessage "Verifying IIS installation (Attempt $i of $RetryCount)..."
        
        # Check Windows Feature status
        $iisFeature = Get-WindowsFeature -Name Web-Server
        if (-not $iisFeature.Installed) {
            Write-LogMessage "IIS feature not installed, installing now..."
            Install-WindowsFeature -Name Web-Server -IncludeManagementTools
            Start-Sleep -Seconds $RetryDelay
            continue
        }
        
        # Check service status
        $iisService = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
        if (-not $iisService) {
            Write-LogMessage "W3SVC service not found, waiting..."
            Start-Sleep -Seconds $RetryDelay
            continue
        }
        
        # Try to start the service
        try {
            if ($iisService.Status -ne 'Running') {
                Start-Service W3SVC -ErrorAction Stop
                Start-Sleep -Seconds 2
            }
            return $true
        } catch {
            Write-LogMessage "Could not start W3SVC service, retrying..."
            Start-Sleep -Seconds $RetryDelay
            continue
        }
    }
    
    Write-LogError "IIS installation verification failed after $RetryCount attempts"
    return $false
}

function Initialize-RequiredModules {
    $requiredModules = @(
        @{
            Name = "WebAdministration"
            InstallCommand = { Install-WindowsFeature Web-Scripting-Tools }
        }
    )

    foreach ($module in $requiredModules) {
        Write-LogMessage "Checking for $($module.Name) module..."
        
        if (-not (Get-Module -ListAvailable -Name $module.Name)) {
            Write-LogMessage "Module $($module.Name) not found. Installing..."
            try {
                & $module.InstallCommand
                Start-Sleep -Seconds 2  # Give Windows time to register the feature
            }
            catch {
                Write-LogError "Failed to install $($module.Name) module: $_"
                return $false
            }
        }

        try {
            Import-Module $module.Name -ErrorAction Stop
            Write-LogMessage "Successfully loaded $($module.Name) module"
        }
        catch {
            Write-LogError "Failed to import $($module.Name) module: $_"
            return $false
        }
    }
    return $true
}

function Test-Prerequisites {
    Write-LogMessage "Checking prerequisites..."

    # Check Windows features
    $requiredFeatures = @(
        'Web-Server',
        'Web-Common-Http',
        'Web-Default-Doc',
        'Web-Dir-Browsing',
        'Web-Http-Errors',
        'Web-Static-Content',
        'Web-Http-Logging',
        'Web-Stat-Compression',
        'Web-Filtering',
        'Web-Mgmt-Console',
        'Web-Asp-Net45',
        'Web-ISAPI-Ext',
        'Web-ISAPI-Filter',
        'Web-Net-Ext45'
    )

    $missingFeatures = @()
    foreach ($feature in $requiredFeatures) {
        $state = Get-WindowsFeature -Name $feature
        if (-not $state.Installed) {
            $missingFeatures += $feature
        }
    }

    if ($missingFeatures.Count -gt 0) {
        Write-LogMessage "Installing missing Windows features: $($missingFeatures -join ', ')"
        try {
            Install-WindowsFeature -Name $missingFeatures -IncludeManagementTools
            Start-Sleep -Seconds 5  # Give Windows time to complete installation
        }
        catch {
            Write-LogError "Failed to install Windows features: $_"
            return $false
        }
    }

    # Check .NET Framework
    try {
        Write-LogMessage "Checking .NET Framework..."
        
        # Method 1: Registry check
        $netVersion = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
        if ($netVersion.Release -ge 461808) {
            $version = "4.7.2 or later"
        } elseif ($netVersion.Release -ge 461308) {
            $version = "4.7.1"
        } elseif ($netVersion.Release -ge 460798) {
            $version = "4.7"
        } elseif ($netVersion.Release -ge 394802) {
            $version = "4.6.2"
        } elseif ($netVersion.Release -ge 394254) {
            $version = "4.6.1"
        } else {
            $version = "4.6 or earlier"
        }

        if ($netVersion) {
            Write-LogMessage ".NET Framework Version: $version (Release: $($netVersion.Release))"
        } else {
            # Method 2: PowerShell check
            $psVersion = $PSVersionTable.CLRVersion
            if ($psVersion) {
                Write-LogMessage ".NET Framework Version: $($psVersion.ToString())"
            } else {
                throw "Could not detect .NET Framework version"
            }
        }

        # Check if ASP.NET is installed
        $aspNet = Get-WindowsFeature -Name Web-Asp-Net45
        if (-not $aspNet.Installed) {
            Write-LogMessage "Installing ASP.NET 4.5..."
            Install-WindowsFeature -Name Web-Asp-Net45 -ErrorAction Stop
            Start-Sleep -Seconds 5
        }

        return $true
    }
    catch {
        Write-LogError "Failed to verify .NET Framework: $_"
        Write-LogMessage "Attempting to install/repair .NET Framework..."
        
        try {
            # Install or repair .NET Framework
            $features = @(
                'NET-Framework-45-Features',
                'NET-Framework-45-Core',
                'NET-Framework-45-ASPNET',
                'NET-WCF-Services45'
            )
            
            foreach ($feature in $features) {
                $state = Get-WindowsFeature -Name $feature
                if (-not $state.Installed) {
                    Write-LogMessage "Installing $feature..."
                    Install-WindowsFeature -Name $feature -ErrorAction Stop
                }
            }
            
            Write-LogMessage ".NET Framework components installed successfully"
            return $true
        }
        catch {
            Write-LogError "Failed to install .NET Framework components: $_"
            return $false
        }
    }

    # Check disk space (need at least 1GB free)
    try {
        $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
        $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        if ($freeSpaceGB -lt 1) {
            Write-LogError "Insufficient disk space. Need at least 1GB, found $freeSpaceGB GB"
            return $false
        }
        Write-LogMessage "Disk space check passed. Free space: $freeSpaceGB GB"
    }
    catch {
        Write-LogError "Failed to check disk space: $_"
        return $false
    }

    return $true
}

function Initialize-IISProvider {
    try {
        Write-LogMessage "Initializing IIS Provider..."
        
        # First ensure IIS is installed
        $iisFeature = Get-WindowsFeature -Name Web-Server
        if (-not $iisFeature.Installed) {
            Write-LogMessage "Installing IIS Web-Server feature..."
            Install-WindowsFeature -Name Web-Server -IncludeManagementTools
            Start-Sleep -Seconds 5
        }

        # Install Web-Scripting-Tools if not present
        $scriptingTools = Get-WindowsFeature -Name Web-Scripting-Tools
        if (-not $scriptingTools.Installed) {
            Write-LogMessage "Installing Web-Scripting-Tools..."
            Install-WindowsFeature -Name Web-Scripting-Tools
            Start-Sleep -Seconds 5
        }

        # Try to import module
        if (-not (Get-Module -Name WebAdministration)) {
            Write-LogMessage "Importing WebAdministration module..."
            Import-Module WebAdministration -SkipEditionCheck -ErrorAction Stop
            Start-Sleep -Seconds 2
        }

        # Verify IIS provider is available
        $provider = Get-PSProvider -PSProvider WebAdministration -ErrorAction SilentlyContinue
        if (-not $provider) {
            throw "WebAdministration provider not available after module import"
        }

        # Create or verify IIS drive
        if (-not (Get-PSDrive -Name IIS -ErrorAction SilentlyContinue)) {
            Write-LogMessage "Creating IIS PSDrive..."
            New-PSDrive -Name IIS -PSProvider WebAdministration -Root "IIS:" -Scope Global -ErrorAction Stop | Out-Null
            Start-Sleep -Seconds 2
        }

        # Final verification
        if (-not (Test-Path "IIS:\")) {
            throw "IIS drive not accessible after creation"
        }

        Write-LogMessage "IIS Provider initialized successfully"
        return $true
    }
    catch {
        Write-LogError "Failed to initialize IIS Provider: $_"
        Write-LogMessage "Attempting alternative initialization method..."
        
        try {
            # Alternative method using PowerShell compatibility
            if (-not (Get-Module -Name WebAdministration)) {
                Import-Module WebAdministration -UseWindowsPowerShell -ErrorAction Stop
            }
            
            if (-not (Get-PSDrive -Name IIS -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name IIS -PSProvider WebAdministration -Root "IIS:" -Scope Global -ErrorAction Stop | Out-Null
            }

            if (Test-Path "IIS:\") {
                Write-LogMessage "IIS Provider initialized using compatibility mode"
                return $true
            }
            throw "IIS drive still not accessible"
        }
        catch {
            Write-LogError "All IIS Provider initialization attempts failed: $_"
            return $false
        }
    }
}

function New-IISWebsite {
    param(
        [string]$SiteName,
        [string]$PhysicalPath,
        [string]$AppPoolName,
        [int]$Port = 80
    )
    
    try {
        # Ensure IIS drive is available
        if (-not (Test-Path "IIS:\")) {
            if (-not (Initialize-IISProvider)) {
                throw "Cannot proceed without IIS Provider"
            }
        }
        
        # Stop Default Web Site and remove its bindings
        Write-LogMessage "Stopping Default Web Site and removing bindings..."
        if (Test-Path "IIS:\Sites\Default Web Site") {
            Stop-Website -Name "Default Web Site" -ErrorAction SilentlyContinue
            Remove-Website -Name "Default Web Site" -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
        
        # Check and remove any existing bindings on port 80
        $existingBindings = Get-WebBinding -Port $Port
        foreach ($binding in $existingBindings) {
            Write-LogMessage "Removing existing binding: $($binding.bindingInformation)"
            $binding | Remove-WebBinding -ErrorAction SilentlyContinue
        }
        
        # Stop and remove existing website if it exists
        if (Test-Path "IIS:\Sites\$SiteName") {
            Write-LogMessage "Removing existing website '$SiteName'..."
            try {
                Stop-Website -Name $SiteName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Get-WebBinding -Name $SiteName | Remove-WebBinding -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Remove-Website -Name $SiteName -ErrorAction Stop
                Start-Sleep -Seconds 2
            } catch {
                Write-LogError "Failed to remove existing website: $_"
            }
        }
        
        # Create and configure application pool
        Write-LogMessage "Configuring application pool '$AppPoolName'..."
        if (Test-Path "IIS:\AppPools\$AppPoolName") {
            try {
                Stop-WebAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Remove-WebAppPool -Name $AppPoolName -ErrorAction Stop
                Start-Sleep -Seconds 2
            } catch {
                Write-LogError "Failed to remove existing app pool: $_"
            }
        }
        
        $appPool = New-WebAppPool -Name $AppPoolName -Force
        Start-Sleep -Seconds 2
        
        # Configure app pool with retry
        $retryCount = 3
        $success = $false
        
        for ($i = 1; $i -le $retryCount; $i++) {
            try {
                Set-ItemProperty "IIS:\AppPools\$AppPoolName" -Name "managedRuntimeVersion" -Value "v4.0"
                Set-ItemProperty "IIS:\AppPools\$AppPoolName" -Name "managedPipelineMode" -Value "Integrated"
                $success = $true
                break
            } catch {
                Write-LogMessage "Attempt $i of $retryCount to configure app pool failed: $_"
                Start-Sleep -Seconds 2
            }
        }
        
        if (-not $success) {
            throw "Failed to configure application pool after $retryCount attempts"
        }
        
        # Create website with retry
        Write-LogMessage "Creating new website..."
        $website = $null
        $success = $false
        
        # Ensure no websites are using port 80
        Get-Website | Where-Object { $_.State -eq 'Started' } | Stop-Website
        Start-Sleep -Seconds 2
        
        for ($i = 1; $i -le $retryCount; $i++) {
            try {
                # Remove any lingering bindings
                Get-WebBinding -Port $Port | Remove-WebBinding
                Start-Sleep -Seconds 2
                
                $website = New-Website -Name $SiteName `
                                     -PhysicalPath $PhysicalPath `
                                     -ApplicationPool $AppPoolName `
                                     -Port $Port `
                                     -Force
                if ($website) {
                    $success = $true
                    break
                }
            } catch {
                Write-LogMessage "Attempt $i of $retryCount to create website failed: $_"
                Start-Sleep -Seconds 2
            }
        }
        
        if (-not $success) {
            throw "Failed to create website after $retryCount attempts"
        }
        
        # Start website with retry
        Write-LogMessage "Starting website..."
        $success = $false
        for ($i = 1; $i -le $retryCount; $i++) {
            try {
                Start-Sleep -Seconds 2
                Start-Website -Name $SiteName -ErrorAction Stop
                $success = $true
                break
            } catch {
                Write-LogMessage "Attempt $i of $retryCount to start website failed: $_"
                if ($_.Exception.Message -match "file already exists") {
                    Write-LogMessage "Attempting to fix binding issue..."
                    try {
                        Get-WebBinding -Port $Port | Remove-WebBinding
                        Start-Sleep -Seconds 2
                        New-WebBinding -Name $SiteName -Protocol "http" -Port $Port -IPAddress "*"
                    } catch {
                        Write-LogError "Failed to update binding: $_"
                    }
                }
                Start-Sleep -Seconds 2
            }
        }
        
        if (-not $success) {
            throw "Failed to start website after $retryCount attempts"
        }
        
        Write-LogMessage "Website created and started successfully"
        return $true
    } catch {
        Write-LogError "Failed to create IIS website: $_"
        return $false
    }
}

# Main script execution
try {
    Write-LogMessage "Starting IIS installation and configuration..."

    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "This script must be run as Administrator"
    }

    if (-not (Initialize-RequiredModules)) {
        throw "Failed to initialize required PowerShell modules"
    }
    if (-not (Initialize-IISProvider)) {
        throw "Failed to initialize IIS Provider"
    }
    if (-not (Test-Prerequisites)) {
        throw "Prerequisites check failed"
    }

    Write-LogMessage "Stopping IIS if running..."
    Stop-Service -Name W3SVC -Force -ErrorAction SilentlyContinue

    Write-LogMessage "Installing IIS and required features..."
    try {
        $installResult = Install-WindowsFeature -Name Web-Server, `
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
        
        if (-not $installResult.Success) {
            throw "Failed to install IIS features. Exit code: $($installResult.ExitCode)"
        }
        
        Write-LogMessage "IIS features installation completed. Verifying..."
        Start-Sleep -Seconds 5  # Give Windows time to initialize services
        
    } catch {
        Write-LogError "Failed to install IIS features: $_"
        exit 1
    }

    # Verify IIS installation with retry
    if (-not (Test-IISInstallation -RetryCount 3 -RetryDelay 5)) {
        Write-LogError "IIS installation verification failed. Please try the following:"
        Write-Host "1. Run 'sfc /scannow' to check system files"
        Write-Host "2. Check Windows Update for pending updates"
        Write-Host "3. Review Event Viewer for errors"
        Write-Host "4. Try rebooting and running the script again"
        exit 1
    }

    Write-LogMessage "IIS installation verified successfully"

    # Get web shells directory from user
    $webShellsPath = "C:\webshells"
    if (-not (Test-Path $webShellsPath)) {
        New-Item -ItemType Directory -Path $webShellsPath -Force | Out-Null
        Write-LogMessage "Created directory: $webShellsPath"
    }

    # Add this line to save the path to a file
    $webRootFile = Join-Path $PSScriptRoot "webrootpath.txt"
    $webShellsPath | Out-File -FilePath $webRootFile -Encoding utf8 -Force

    # Create new IIS website
    Write-LogMessage "Configuring IIS website..."
    if (-not (New-IISWebsite -SiteName "WebShells" `
                            -PhysicalPath $webShellsPath `
                            -AppPoolName "WebShellsPool" `
                            -Port 80)) {
        throw "Failed to configure IIS website"
    }

    # Start IIS
    Write-LogMessage "Starting IIS..."
    try {
        Start-Service W3SVC
    } catch {
        Write-LogError "Failed to start IIS: $_"
        exit 1
    }

    # Create test file
    $testFile = Join-Path $webShellsPath "test.aspx"
@"
<%@ Page Language="C#" %>
<!DOCTYPE html>
<html>
<head><title>IIS Test</title></head>
<body>
<h1>IIS is working!</h1>
<p>Server Time: <%= DateTime.Now %></p>
</body>
</html>
"@ | Out-File -FilePath $testFile -Encoding UTF8

    Write-LogMessage "Installation completed successfully!"

    # Display connection information
    Write-Host "`n=== IIS Service Status ===" -ForegroundColor Green
    Get-Service W3SVC | Format-Table Name, Status, DisplayName

    Write-Host "`n=== Connection Information ===" -ForegroundColor Green
    Write-Host "Local URL: http://localhost/"

    Write-Host "`nAvailable IP addresses to connect to:" -ForegroundColor Green
    Get-AllIPAddresses | ForEach-Object {
        Write-Host "http://$($_.IPAddress)/"
    }

    Write-Host "`nPublic IP address:" -ForegroundColor Green
    try {
        $publicIP = (Invoke-WebRequest -Uri "http://ifconfig.me/ip" -UseBasicParsing).Content
        Write-Host "http://$publicIP/"
    } catch {
        Write-Host "Could not determine public IP address"
    }

    Write-Host "`n=== Important File Locations ===" -ForegroundColor Green
    Write-Host "Web root directory: $webShellsPath"
    Write-Host "IIS configuration: %SystemRoot%\System32\inetsrv\config\"
    Write-Host "IIS logs: %SystemRoot%\System32\LogFiles\"

    Write-Host "`n=== Files available in $webShellsPath ===" -ForegroundColor Green
    Get-ChildItem $webShellsPath | Format-Table Name, Length, LastWriteTime

    Write-Host "`n=== Verification Steps ===" -ForegroundColor Green
    Write-Host "1. IIS Status: $((Get-Service W3SVC).Status)"
    Write-Host "2. .NET Version: $([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)"
    Write-Host "3. Directory Permissions: "
    Get-Acl $webShellsPath | Format-List

    Write-LogMessage "Setup complete! You can now access your web shells through IIS."

    return $webShellsPath
} catch {
    Write-LogError "Critical error during setup: $_"
    Write-LogError "Stack Trace: $($_.ScriptStackTrace)"
    exit 1
} 