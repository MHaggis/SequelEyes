<#
.SYNOPSIS
    SQL Server and IIS Integration Setup Script

.DESCRIPTION
    This script configures SQL Server and IIS integration by:
    - Creating and configuring a sample database
    - Setting up IIS with SQL Server authentication
    - Establishing connectivity between IIS and SQL Server
    - Creating test data and verifying the setup

.PARAMETER SQLServer
    The SQL Server instance to configure (default: localhost)

.PARAMETER Database  
    The name of the sample database to create (default: CustomerDB)

.PARAMETER SQLPassword
    SQL Server sa account password (required)

.EXAMPLE
    # Basic usage with default SQL instance
    .\sql_iis_setup.ps1

.EXAMPLE
    # Configure named SQL instance with custom database
    .\sql_iis_setup.ps1

.NOTES
    File Name      : sql_iis_setup.ps1
    Author         : The Haag
    Prerequisite   : PowerShell 5.1 or later
                    SQL Server instance
                    IIS installed and configured
                    Administrative privileges

.LINK
    https://github.com/MHaggis/SequelEyes
#>

$ErrorActionPreference = "Stop"


#Requires -RunAsAdministrator

function Write-LogMessage {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Write-LogError {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $Message" -ForegroundColor Red
}

function Initialize-SampleDatabase {
    param (
        [string]$SQLServer = "localhost",
        [string]$Database = "CustomerDB",
        [string]$SQLPassword
    )
    
    try {
        Write-LogMessage "Creating sample database..."
        $conn = New-Object System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = "Server=$SQLServer;Database=master;User ID=sa;Password=$SQLPassword"
        $conn.Open()

        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.Connection = $conn
        $cmd.CommandText = "IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = '$Database') CREATE DATABASE $Database"
        $cmd.ExecuteNonQuery()
        $conn.Close()

        $conn.ConnectionString = "Server=$SQLServer;Database=$Database;User ID=sa;Password=$SQLPassword"
        $conn.Open()
        $cmd.CommandText = @"
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Customers' and xtype='U')
        CREATE TABLE Customers (
            CustomerID INT PRIMARY KEY IDENTITY(1,1),
            FirstName NVARCHAR(50),
            LastName NVARCHAR(50),
            Email NVARCHAR(100),
            Phone NVARCHAR(20),
            Address NVARCHAR(200)
        )
"@
        $cmd.ExecuteNonQuery()

        $cmd.CommandText = @"
        IF NOT EXISTS (SELECT * FROM Customers)
        BEGIN
            INSERT INTO Customers (FirstName, LastName, Email, Phone, Address) VALUES
            ('John', 'Doe', 'john.doe@email.com', '555-0100', '123 Main St'),
            ('Jane', 'Smith', 'jane.smith@email.com', '555-0101', '456 Oak Ave'),
            ('Bob', 'Johnson', 'bob.j@email.com', '555-0102', '789 Pine Rd'),
            ('Alice', 'Brown', 'alice.b@email.com', '555-0103', '321 Elm St'),
            ('Charlie', 'Wilson', 'charlie.w@email.com', '555-0104', '654 Maple Dr')
        END
"@
        $cmd.ExecuteNonQuery()
        $conn.Close()
        
        Write-LogMessage "Sample database created and populated successfully"
        return $true
    }
    catch {
        Write-LogError "Failed to initialize database: $_"
        return $false
    }
}

function New-QueryPage {
    param([string]$WebRoot)
    
    try {
        $aspxContent = @"
<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<!DOCTYPE html>
<html>
<head>
    <title>SQL Testing Interface</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 20px; background: #1e1e1e; color: #fff; }
        .container { max-width: 1400px; margin: 0 auto; }
        .header { background: #2d2d2d; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .query-box { width: 100%; height: 200px; background: #2d2d2d; color: #fff; font-family: 'Consolas', monospace; 
                    padding: 15px; margin: 10px 0; border: 1px solid #3d3d3d; border-radius: 4px; font-size: 14px; }
        .button { background: #0078d4; color: white; padding: 10px 20px; border: none; border-radius: 4px; 
                 cursor: pointer; font-size: 14px; }
        .button:hover { background: #106ebe; }
        .results { margin-top: 20px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; background: #2d2d2d; }
        th { background: #3d3d3d; padding: 10px; text-align: left; }
        td { padding: 8px; border: 1px solid #3d3d3d; }
        .error { color: #ff6b6b; background: #2d2d2d; padding: 10px; border-radius: 4px; margin: 10px 0; }
        .success { color: #69db7c; background: #2d2d2d; padding: 10px; border-radius: 4px; margin: 10px 0; }
        .test-queries { background: #2d2d2d; padding: 20px; border-radius: 8px; margin: 20px 0; }
        .test-queries h3 { margin-top: 0; color: #0078d4; }
        .query-example { background: #3d3d3d; padding: 10px; border-radius: 4px; margin: 10px 0; 
                        font-family: 'Consolas', monospace; cursor: pointer; }
        .query-example:hover { background: #4d4d4d; }
        .copy-button { float: right; background: #0078d4; border: none; color: white; padding: 2px 8px; 
                      border-radius: 3px; cursor: pointer; }
        .copy-button:hover { background: #106ebe; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>SQL Testing Interface</h1>
        </div>
        <form runat="server">
            <asp:TextBox runat="server" ID="QueryInput" TextMode="MultiLine" CssClass="query-box" />
            <br />
            <asp:Button runat="server" Text="Execute Query" OnClick="ExecuteQuery" CssClass="button" />
            <div id="results" runat="server" class="results"></div>
        </form>

        <div class="test-queries">
            <h3>Test Queries</h3>
            
            <h4>1. Basic Tests</h4>
            <div class="query-example" onclick="copyToQueryBox(this)">
                SELECT @@VERSION AS Version;
                SELECT GETDATE() AS ServerTime;
                SELECT SYSTEM_USER AS CurrentUser, SERVERPROPERTY('MachineName') AS MachineName;
                <button class="copy-button">Copy</button>
            </div>

            <h4>2. System Info</h4>
            <div class="query-example" onclick="copyToQueryBox(this)">
                SELECT 
                    SERVERPROPERTY('ProductVersion') AS ProductVersion,
                    SERVERPROPERTY('Edition') AS Edition,
                    SERVERPROPERTY('ProductLevel') AS ProductLevel,
                    SERVERPROPERTY('Collation') AS Collation;

                SELECT 
                    DB_NAME(database_id) AS DatabaseName,
                    CAST(SUM(size) * 8. / 1024 AS DECIMAL(8,2)) AS SizeMB
                FROM sys.master_files
                GROUP BY database_id
                ORDER BY SUM(size) DESC;
                <button class="copy-button">Copy</button>
            </div>

            <h4>3. Security Checks</h4>
            <div class="query-example" onclick="copyToQueryBox(this)">
                SELECT name, type_desc, create_date 
                FROM sys.server_principals 
                WHERE type_desc NOT LIKE '%CERTIFICATE%';

                SELECT * FROM fn_my_permissions(NULL, 'SERVER');
                <button class="copy-button">Copy</button>
            </div>

            <h4>4. Enable xp_cmdshell</h4>
            <div class="query-example" onclick="copyToQueryBox(this)">
                sp_configure 'show advanced options', 1;
                RECONFIGURE;
                
                sp_configure 'xp_cmdshell', 1;
                RECONFIGURE;
                
                EXEC xp_cmdshell 'whoami';
                <button class="copy-button">Copy</button>
            </div>

            <h4>5. System Enumeration</h4>
            <div class="query-example" onclick="copyToQueryBox(this)">
                SELECT * FROM sys.objects;
                SELECT @@SERVERNAME, @@VERSION;
                <button class="copy-button">Copy</button>
            </div>

            <h4>6. Data Export Tests</h4>
            <div class="query-example" onclick="copyToQueryBox(this)">
                SELECT name, type FROM sys.objects FOR XML AUTO;
                SELECT name, type FROM sys.objects FOR JSON AUTO;
                SELECT name, type INTO #temp FROM sys.objects;
                <button class="copy-button">Copy</button>
            </div>
        </div>
    </div>

    <script type="text/javascript">
        function copyToQueryBox(element) {
            var queryBox = document.getElementById('<%= QueryInput.ClientID %>');
            queryBox.value = element.innerText.replace('Copy', '').trim();
            queryBox.focus();
        }
    </script>

    <script runat="server">
        protected void ExecuteQuery(object sender, EventArgs e)
        {
            try {
                string[] queries = QueryInput.Text.Split(new[] { "GO", ";" }, StringSplitOptions.RemoveEmptyEntries);
                StringBuilder output = new StringBuilder();
                
                using (SqlConnection conn = new SqlConnection("Server=localhost;Database=master;User ID=sa;Password=ComplexPass123!;MultipleActiveResultSets=true"))
                {
                    conn.Open();
                    
                    foreach (string query in queries)
                    {
                        if (string.IsNullOrWhiteSpace(query)) continue;
                        
                        using (SqlCommand cmd = new SqlCommand(query.Trim(), conn))
                        {
                            cmd.CommandTimeout = 60;
                            
                            try {
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    do {
                                        if (reader.HasRows)
                                        {
                                            output.Append("<table>");
                                            
                                            // Headers
                                            output.Append("<tr>");
                                            for (int i = 0; i < reader.FieldCount; i++)
                                            {
                                                output.Append(string.Format("<th>{0}</th>", reader.GetName(i)));
                                            }
                                            output.Append("</tr>");
                                            
                                            // Data
                                            while (reader.Read())
                                            {
                                                output.Append("<tr>");
                                                for (int i = 0; i < reader.FieldCount; i++)
                                                {
                                                    output.Append(string.Format("<td>{0}</td>", 
                                                        Server.HtmlEncode(reader[i].ToString())));
                                                }
                                                output.Append("</tr>");
                                            }
                                            
                                            output.Append("</table>");
                                        }
                                    } while (reader.NextResult());
                                    
                                    if (!reader.HasRows)
                                    {
                                        output.Append("<div class='success'>Command executed successfully.</div>");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                output.Append(string.Format("<div class='error'>Error executing query: {0}</div>", 
                                    Server.HtmlEncode(ex.Message)));
                            }
                        }
                    }
                }
                
                results.InnerHtml = output.ToString();
            }
            catch (Exception ex)
            {
                results.InnerHtml = string.Format("<div class='error'>Error: {0}</div>", 
                    Server.HtmlEncode(ex.Message));
            }
        }
    </script>
</body>
</html>
"@

        $aspxPath = Join-Path $WebRoot "query.aspx"
        Set-Content -Path $aspxPath -Value $aspxContent
        Write-LogMessage "Query page created at: $aspxPath"
        return $true
    }
    catch {
        Write-LogError "Failed to create query page: $_"
        return $false
    }
}

function Test-Installation {
    param(
        [string]$WebRoot,
        [string]$SQLPassword
    )
    
    try {
        Write-LogMessage "Validating installation..."
        
        $iisService = Get-Service -Name W3SVC -ErrorAction Stop
        if ($iisService.Status -ne 'Running') {
            throw "IIS service is not running"
        }
        
        $sqlService = Get-Service -Name MSSQLSERVER -ErrorAction Stop
        if ($sqlService.Status -ne 'Running') {
            throw "SQL Server service is not running"
        }
        
        $conn = New-Object System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = "Server=localhost;Database=CustomerDB;User ID=sa;Password=$SQLPassword"
        $conn.Open()
        $conn.Close()
        
        if (-not (Test-Path (Join-Path $WebRoot "query.aspx"))) {
            throw "Query page not found"
        }
        
        Write-LogMessage "Installation validation completed successfully"
        return $true
    }
    catch {
        Write-LogError "Installation validation failed: $_"
        return $false
    }
}

function Write-Banner {
    $banner = @"
+------------------------------------------------------------------------+
 #####                                     #######                    
#     # ######  ####  #    # ###### #      #       #   # ######  #### 
#       #      #    # #    # #      #      #        # #  #      #     
 #####  #####  #    # #    # #####  #      #####     #   #####   #### 
      # #      #  # # #    # #      #      #         #   #           #
#     # #      #   #  #    # #      #      #         #   #      #    #
 #####  ######  ### #  ####  ###### ###### #######   #   ######  ####                  
+------------------------------------------------------------------------+
"@
    Write-Host $banner -ForegroundColor Cyan
}

function Write-SectionHeader {
    param([string]$Title)
    Write-Host "`n=== $Title ===" -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Message)
    Write-Host "[SUCCESS] $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Summary {
    param(
        [Parameter(Mandatory=$true)]
        [string]$WebRoot,
        [Parameter(Mandatory=$true)]
        [hashtable]$ServiceStatus,
        [Parameter(Mandatory=$true)]
        [string]$SqlVersion,
        [Parameter(Mandatory=$true)]
        [string]$NetVersion
    )
    
    Clear-Host
    Write-Banner
    
    Write-SectionHeader "Installation Summary"
    Write-Host @"

Web Interface
------------
Local URL........: http://localhost/query.aspx
IP Address.......: http://$((Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.InterfaceAlias -like "*Ethernet*"}).IPAddress)/query.aspx
Web Root.........: $WebRoot

Service Status
-------------
IIS..............: $($ServiceStatus.IIS)
SQL Server.......: $($ServiceStatus.SQL)

Versions
--------
SQL Server.......: $SqlVersion
.NET Framework...: $NetVersion

File Locations
-------------
Web Root.........: $WebRoot
IIS Config.......: %SystemRoot%\System32\inetsrv\config\
IIS Logs.........: %SystemRoot%\System32\LogFiles\

Security
--------
IIS App Pool.....: WebShellsPool (Identity: ApplicationPoolIdentity)
SQL Login........: sa (Password: ComplexPass123!)

Available Test Queries
--------------------
1. Basic Tests (Version, Time, User)
2. System Information
3. Security Checks
4. xp_cmdshell Configuration
5. System Enumeration
6. Data Export Tests

Next Steps
---------
1. Open http://localhost/query.aspx in your browser
2. Use the pre-configured test queries
3. Start your security testing

"@
}

try {
    Write-Banner
    Write-Info "Starting SQL Server and IIS setup..."
    
    Write-Info "Installing IIS..."
    $iisScript = Join-Path $PSScriptRoot "install_iis_aspnet.ps1"
    if (-not (Test-Path $iisScript)) {
        throw "IIS installation script not found at: $iisScript"
    }

    # Run the IIS script
    & $iisScript

    # Get the actual path from IIS
    try {
        $website = Get-Website -Name "WebShells"
        $webRoot = $website.PhysicalPath
        
        if (-not $webRoot) {
            throw "Could not get website physical path"
        }
        
        Write-LogMessage "Using web root path from IIS: $webRoot"
    } catch {
        Write-LogMessage "Failed to get web root path from IIS. Please enter it manually:"
        $webRoot = Read-Host "Web root path"
    }

    # Verify the path exists
    if (-not (Test-Path $webRoot)) {
        throw "Web root path does not exist: $webRoot"
    }

    $maxAttempts = 3
    $attempt = 1
    $success = $false
    
    while ($attempt -le $maxAttempts) {
        Write-LogMessage "Verifying IIS installation (Attempt $attempt of $maxAttempts)..."
        
        try {
            $iisService = Get-Service -Name W3SVC -ErrorAction Stop
            $website = Get-Website -Name "WebShells" -ErrorAction Stop
            
            if ($iisService.Status -eq 'Running' -and $website) {
                $success = $true
                break
            }
        } catch {
            Write-LogMessage "Verification attempt $attempt failed: $_"
        }
        
        Start-Sleep -Seconds 5
        $attempt++
    }
    
    if (-not $success) {
        throw "IIS installation verification failed after $maxAttempts attempts"
    }
    
    if (Test-Path "C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER") {
        $reinstall = Read-Host "SQL Server is already installed. Do you want to proceed with database setup? (Y/N)"
        if ($reinstall -ne 'Y' -and $reinstall -ne 'y') {
            throw "Installation cancelled by user"
        }
    } else {
        Write-LogMessage "Installing SQL Server..."
        $sqlScript = Join-Path (Split-Path $PSScriptRoot -Parent) "SQLSSTT\install-SQL.ps1"
        if (-not (Test-Path $sqlScript)) {
            throw "SQL installation script not found at: $sqlScript"
        }
        & $sqlScript
        if ($LASTEXITCODE -ne 0) {
            throw "SQL Server installation failed"
        }
    }
    
    Write-LogMessage "Proceeding with database setup..."
    
    if (-not (Initialize-SampleDatabase -SQLPassword "ComplexPass123!")) {
        throw "Failed to initialize sample database"
    }
    
    if (-not (New-QueryPage -WebRoot $webRoot)) {
        throw "Failed to create query page"
    }
    
    if (-not (Test-Installation -WebRoot $webRoot -SQLPassword "ComplexPass123!")) {
        throw "Installation validation failed"
    }
    
    $serviceStatus = @{
        IIS = (Get-Service W3SVC).Status
        SQL = (Get-Service MSSQLSERVER).Status
    }
    
    try {
        $sqlVersion = Invoke-Sqlcmd -Query "SELECT @@VERSION" `
            -ServerInstance localhost `
            -Username sa `
            -Password "ComplexPass123!" `
            -TrustServerCertificate $true | 
            Select-Object -ExpandProperty Column1
    }
    catch {
        $sqlVersion = "Could not retrieve SQL version"
    }
    
    $netVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release
    
    Write-Summary `
        -WebRoot $webRoot `
        -ServiceStatus $serviceStatus `
        -SqlVersion $sqlVersion `
        -NetVersion $netVersion
    
    Write-Success "Setup completed successfully!"
} catch {
    Write-Host "Setup failed: $_" -ForegroundColor Red
    exit 1
}
