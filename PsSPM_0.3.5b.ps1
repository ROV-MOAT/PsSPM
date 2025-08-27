<#
.SYNOPSIS
    PsSPM(ROV-MOAT) - Powershell SNMP Printer Monitoring and Reporting Script
    C# SNMP Library is used - https://github.com/lextudio/sharpsnmplib

.DESCRIPTION
    Checks printer status via SNMP and generates CSV, HTML report with toner levels, counters, and information

.NOTES
    Version: 0.3.5b
    Author: Oleg Ryabinin + AI
    Date: 2025-08-27
    
    MESSAGE:
    Powershell 5+

    CHANGELOG:

    Ver. 0.3.5b
    + Mail send
    + Interface Mode (Console, FullGui, LightGui)
    Optimization
    
    Ver. 0.3.4
    + WPF GUI
    - WF GUI
    + IP range

    Ver. 0.3.3
    + RunspacePool (Test-TcpConnectionParallel)

    Ver. 0.3.2
    Visual changes in HTML report
    + WF GUI
    + CSV Report + Buffer ($CsvBufferSize)
    
    Ver. 0.3.1
    Visual changes in HTML report
    + Write-Log
    + $MaxRetries, $RetryDelayMs (Test-PrinterConnection)

    Ver. 0.3
    Visual changes in HTML report
    For SNMP requests, the Lextm.SharpSnmpLib ver.12.5.6.0 library is used.

    Ver. 0.2
    Visual changes in HTML report
    Changed the Online/Offline check method (Test-Connection replaced with System.Net.Sockets.TcpClient)..

    Ver. 0.1
    Release

.LINK
    PsSPM(ROV-MOAT) - https://github.com/ROV-MOAT/PsSPM
    C# SNMP Library - https://github.com/lextudio/sharpsnmplib
    Download C# SNMP Library - https://www.nuget.org/packages/Lextm.SharpSnmpLib/
#>
param(
    [ValidateSet("Console", "FullGui", "LightGui")]
    [string]$InterfaceMode = "FullGui",
    [string]$ConsoleFile = $PSScriptRoot
)

#[bool]$ShowGUI = $true
[string]$Version = "0.3.5b"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

#REPORT
[bool]$HtmlFileReport = $true    # On/Off HTML Report out file
[bool]$CsvFileReport  = $false     # On/Off CSV Report out file
[int]$CsvBufferSize = 20     # Number of lines in buffer to write to CSV file

#LOG
[bool]$WriteLog = $false   # On/Off Logging Function
[string]$LogDir = "$PSScriptRoot\Log"
[string]$LogFile = "$logDir\PrinterMonitor_$(Get-Date -Format 'yyyyMMddHHmmss').log"

#SNMP
[int]$TimeoutMsUDP = 5000  # Timeout SNMP - Milliseconds (1 sec. = 1000 ms) / 0 = infinity

#TCP
[int]$TCPPort = 80         # TCP port for check link
[int]$TimeoutMsTCP = 500   # Timeout TcpClient - Milliseconds (1 sec. = 1000 ms)
[int]$MaxRetries   = 3     # Retry Count
[int]$RetryDelayMs = 2000  # Milliseconds (1 sec. = 1000 ms)
[int]$TCPThreads = 10      # Ьaximum number of running TCP threads

# Mail
[bool]$MailSend = $false
[string]$MailFrom = "sender@example.com"
[string[]]$MailTo = "recipient@example.com"
[string]$Subject = "PsSPM Report"
[string[]]$CC   # Carbon copy
[string[]]$BCC  # Blind carbon copy
[string]$SmtpServer = "smtp.example.com"
[int]$SmtpPort = 25
[int]$SmtpTimeoutMs = 10000 # Timeout Smtp - Milliseconds (1 sec. = 1000 ms)

#Path
[string]$PrinterOIDPath = "$PSScriptRoot\Lib\PsSPM_oid.psd1"
[string]$ScriptGuiXaml = "$PSScriptRoot\Lib\PsSPM_gui_wpf.ps1"
[string]$HtmlHeader = "$PSScriptRoot\Lib\PsSPM_html.ps1"
[string]$PrinterListPath = "$PSScriptRoot\IP"
[string]$ReportDir ="$PSScriptRoot\Report"
[string]$DllPath = "$PSScriptRoot\Lib\SharpSnmpLib.dll"
[string]$selectedFile = $null
[string]$filenamecsv = $null
[string]$filenamehtml = $null

#region Helper Functions
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir | Out-Null }

# Logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")

    if ($WriteLog) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "$timestamp [$Level] - $Message"
        Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8
        if ($Level -eq "ERROR") { Write-Host $logMessage -ForegroundColor Red }
        elseif ($Level -eq "WARNING") { Write-Host $logMessage -ForegroundColor Yellow }
        else { Write-Host $logMessage -ForegroundColor Green }
    }
}

# Open File Dialog function
function Open-File([string] $initialDirectory){
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    try {
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if ($OpenFileDialog.ShowDialog() -eq 'OK') { return $OpenFileDialog.filename }
    }
    finally { $OpenFileDialog.Dispose() }
}

# Test-TcpConnectionParallel function
function Test-TcpConnectionParallel {
    param(
        [array]$Devices = @(),
        [int]$Port = 80,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 3000,
        [int]$TimeoutMsTCP = 500,
        [int]$Threads = 5
    )
    $start = Get-Date
    # Create RunspacePool
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
    $runspacePool.Open()
    $runspaces = @()

    # ScriptBlock for connection testing
    $scriptBlock = {
        param($Device, $Port, $MaxRetries, $RetryDelayMs, $TimeoutMsTCP)

        function Connect-TcpClient {
            param($HostName, $Port, $MaxRetries, $RetryDelayMs, $TimeoutMsTCP)

            $attempt = 0
            $lastException = $null

            while ($attempt -lt $MaxRetries) {
                $attempt++
                $client = $null

                try {
                    # Create and configure the TcpClient
                    $client = [System.Net.Sockets.TcpClient]::new()
                    $connectTask = $client.ConnectAsync($HostName, $Port)
                    
                    # Wait for the task with timeout
                    $completedTask = [System.Threading.Tasks.Task]::WaitAny($connectTask, [System.Threading.Tasks.Task]::Delay($TimeoutMsTCP))
                    
                    if ($completedTask -eq 0) {
                        # Connection task completed
                        if ($connectTask.Exception) { throw $connectTask.Exception.InnerException }
                        
                        return [PSCustomObject]@{
                            Device = $HostName
                            Port = $Port
                            Connected = $true
                            Message = "Success"
                            Exception = $null
                        }
                    }
                    else {
                        # Timeout occurred
                        throw [System.TimeoutException]::new("Connection attempt timed out after $TimeoutMsTCP ms")
                    }
                }
                catch [System.Exception] {
                    $lastException = $_

                    if ($attempt -lt $MaxRetries) { Start-Sleep -Milliseconds $RetryDelayMs }
                }
                finally { if ($client) { try { $client.Dispose() } catch {} } }
            }

            return [PSCustomObject]@{
                Device = $HostName
                Port = $Port
                Connected = $false
                Message = if ($lastException) { $_.Message } else { "Unknown error" }
                Exception = if ($lastException) { $_ } else { $null }
            }
        }
        Connect-TcpClient -HostName $Device -Port $Port -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs -TimeoutMs $TimeoutMsTCP
    }

    # Create and start runspaces
    foreach ($Device in $Devices) {
        $powershell = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($Device).AddArgument($Port).AddArgument($MaxRetries).AddArgument($RetryDelayMs).AddArgument($TimeoutMsTCP)
        $powershell.RunspacePool = $runspacePool
        
        $runspaceHandle = [PSCustomObject]@{
            PowerShell = $powershell
            Handle = $powershell.BeginInvoke()
            Device = $Device
        }
        $runspaces += $runspaceHandle
    }

    # Wait for completion and collect results
    $total = $runspaces.Count
    Write-Log "All $total tests started, waiting for completion..."
    $completedCount = 0

    while ($runspaces | Where-Object { -not $_.Handle.IsCompleted }) {
        $completedCount = ($runspaces | Where-Object { $_.Handle.IsCompleted }).Count
        $progress = [math]::Round(($completedCount / $total) * 100, 2)
        Write-Log "Progress: $progress% ($completedCount/$total)"
        Start-Sleep -Milliseconds 500
    }

    $results = @()
    foreach ($runspace in $runspaces) {
        try {
            $results += $runspace.PowerShell.EndInvoke($runspace.Handle)
        }
        catch {
            $results += [PSCustomObject]@{
                Device = $runspace.Device
                Port = $Port
                Connected = $false
                Message = $_.Exception.Message
                Exception = $_.Exception
            }
        }
        finally { try { $runspace.PowerShell.Dispose() } catch {} }
    }
    # Clean up
    $runspacePool.Close()
    $runspacePool.Dispose()

    $end = Get-Date
    $duration = $end - $start
    Write-Log "=== TCP test completed in $($duration.TotalSeconds) seconds ==="
    return $results
}

#SNMP function
function Get-SnmpData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [string]$Community = "public",
        [Parameter(Mandatory=$true)]
        [string]$Oid,
        [int]$UDPport = 161,
        [int]$TimeoutMsUDP = 5000
    )

    try {
        # Set up SNMP manager
        $IP = [System.Net.IPAddress]::Parse($Target)
        $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)
        $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
        $null = $vList.Add([Lextm.SharpSnmpLib.Variable]::new([Lextm.SharpSnmpLib.ObjectIdentifier]::new($Oid)))

        $result = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            [Lextm.SharpSnmpLib.OctetString]::new($Community),
            $vList,
            $TimeoutMsUDP
        )
        return [PSCustomObject]@{
            Success = $true
            result = $result.Data.ToString()
        }
    }
    catch { Write-Log "SNMP query failed for $Target (OID: $Oid): $_" -Level "WARNING"
        return [PSCustomObject]@{
            Success = $false
            result = $null
        }
    }
}

function Get-SnmpBulkWalkWithEncoding {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [string]$Community = "public",
        [Parameter(Mandatory=$true)]
        [string]$Oid,
        [int]$MaxRepetitions = 10,
        [int]$UDPport = 161,
        [int]$TimeoutMsUDP = 5000
    )

    $result_out = [System.Collections.Generic.List[PSObject]]::new()
    $encodings = @([System.Text.Encoding]::UTF8, [System.Text.Encoding]::ASCII, [System.Text.Encoding]::GetEncoding(1251))

    try {
        # Set up SNMP manager
        $IP = [System.Net.IPAddress]::Parse($Target)
        $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)
        $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()

        # Perform the walk
        $null = [Lextm.SharpSnmpLib.Messaging.Messenger]::BulkWalk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            [Lextm.SharpSnmpLib.OctetString]::new($Community),
            [Lextm.SharpSnmpLib.OctetString]::Empty,
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new($Oid),
            $vList,
            $TimeoutMsUDP,
            $MaxRepetitions,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree,
            $null,
            $null
        )
        foreach ($result in $vList) {
            $data = $result.Data
            $value = $null

            # Handle OctetString encoding
            if ($data.TypeCode -eq [Lextm.SharpSnmpLib.SnmpType]::OctetString) {
                foreach ($enc in $encodings) {
                    try { $value = $data.ToString($enc); break }
                    catch { $value = $data.ToString() }
                }
                $null = $result_out.Add([Environment]::NewLine+"$value")
            }
        }
        $cleanArray = $result_out.Where( {$_.Trim() -ne ""} )
        return $cleanArray
    }
    catch { Write-Log "SNMP Walk query failed for $Target (OID: $Oid): $_" -Level "WARNING"; return "Error" }
}

function Get-PrinterModelOIDSet {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Model,
        [Parameter(Mandatory=$true)]
        [hashtable]$OIDMapping
    )
    # More flexible model matching with wildcards
    $modelPatterns = @{
        "*333*" = "333"
        "*B60*" = "B60"
        "*B70*" = "B60"  # Same as B60
        "*M40*" = "M40"
        "*B80*" = "B80"
        "*365*" = "365"
        "*594*" = "365"  # Same as 365
        "*462*" = "365"  # Same as 365
        "*C40*" = "C40"
        "*650*" = "660"  # Same as 660
        "*660*" = "660"
        "*651*" = "651"
        "*T79*" = "T79"
        "*332*" = "332"
    }

    foreach ($pattern in $modelPatterns.Keys) { if ($Model -like $pattern) {return $OIDMapping[$modelPatterns[$pattern]]} }
    return $OIDMapping["Default"]
}

function Get-PrinterData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TargetHost,
        [Parameter(Mandatory=$true)]
        [hashtable]$OidMapping,
        [int]$TimeoutMsUDP = 2000,
        [string]$CurrentPrinter = ""
    )

    $results = [System.Collections.Generic.Dictionary[string, object]]::new()

    try {
        # Get printer model to determine OIDs to use
        $model = Get-SnmpData -Target $TargetHost -Oid $OidMapping["Default"].Model -TimeoutMsUDP $TimeoutMsUDP
        if (-not $model) { Write-Log "Failed to get printer model $TargetHost" -Level "WARNING"
            return $null
        }

        Write-Host "$CurrentPrinter : $TargetHost - $($model.result)" -ForegroundColor Green
        Write-Log "Checking printer: $TargetHost - Online"
        
        $printername = Get-SnmpData -Target $TargetHost -Oid $OidMapping["Default"].PName -TimeoutMsUDP $TimeoutMsUDP
        if (-not $printername) { Write-Log "Failed to get printer name $TargetHost" -Level "WARNING"
            return $null
        }

        # Determine OID set based on model
        $Script:oidSet = Get-PrinterModelOIDSet -Model $model.result -OIDMapping $OidMapping
        if (-not $oidSet) { Write-Log "Not OID set for model: $($model.result)" -Level "WARNING"
            return $null
        }

        # Safe add to dictionary
        $results['Model'] = $model.result
        $results['PName'] = $printername.result
        $results['IPAddress'] = $TargetHost

        foreach ($item in $oidSet.GetEnumerator()) {
            if ($item.Name -in @("Display", "Status")) { continue }
            
            if ($item.Value -is [string]) {
                $Name = $item.Name
                try { $Value = Get-SnmpData -Target $TargetHost -Oid $item.Value -TimeoutMsUDP $TimeoutMsUDP
                    if ($value.Success -like 'true' ) { $results[$Name] = $Value.result } else { throw }
                }
                catch { Write-Log "Error Get-SnmpData for $Name : $($_.Exception.Message)" -Level "WARNING"
                    $results[$Name] = "Error"
                }
            }
        }
        return $results
    }
    catch { Write-Log "Error Get-PrinterData for $TargetHost : $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function Format-Value {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Value,
        [string]$Default = "Error"
    )

    if ($null -eq $Value -or $Value -eq "") { return $null }

    if ($Value -like 'Error') { return "<center class='error'>$Default</center>" }

    return "<center>$($Value.ToString())</center>"
}

function Get-TonerPercentage {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Total,
        [Parameter(Mandatory=$true)]
        [object]$Current
    )

    if ($Total -like 'Error' -or $Current -like 'Error') { return "Error" }
    if ($null -eq $Total -or $null -eq $Current -or $Total -eq "" -or $Current -eq "") { return $null }
    
    try {
        $totalValue = [int]$Total
        $currentValue = [int]$Current
        
        if ($totalValue -le 0 -or $currentValue -le 0) { return 0 }
        
        $percentage = [math]::Round(($currentValue / $totalValue) * 100)
        return $percentage
    }
    catch { return $null }
}
function Update-TonerLevels {
    param([object]$PrinterData)

    $TonerLevels =@{}
    $cartridgeConfig = @{
        'TK'  = @('TonerKTotal', 'TonerKCurrent')
        'DKU' = @('DrumKUTotal', 'DrumKUCurrent')
        'DK'  = @('DrumKTotal', 'DrumKCurrent')
        'TC'  = @('TonerCTotal', 'TonerCCurrent')
        'TM'  = @('TonerMTotal', 'TonerMCurrent')
        'TY'  = @('TonerYTotal', 'TonerYCurrent')
    }
    
    foreach ($key in $cartridgeConfig.Keys) {
        $totalKey = $cartridgeConfig[$key][0]
        $currentKey = $cartridgeConfig[$key][1]
        
        if ($PrinterData.ContainsKey($totalKey) -and $PrinterData.ContainsKey($currentKey)) {
            $TonerLevels[$key] = Get-TonerPercentage -Total $PrinterData[$totalKey] -Current $PrinterData[$currentKey]
        }
    }
    return $TonerLevels
}

function Format-TonerLevel {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Level
    )

    if ($null -eq $Level -or $Level -like "") { return $null }
    if ($Level -like 'Error') { return "<div class='container'><center><span class='error'>Error</span></center></div>" }

    try { $tonerValue = [int]$Level } catch { return $null }

    switch ($tonerValue) {
        { $_ -gt 49 } { return "<div class='container'><center><span class='toner-high'>$Level%</span></center></div>" }
        { $_ -gt 10 -and $_ -le 49 } { return "<div class='container'><center><span class='toner-medium'>$Level%</span></center></div>" }
        { $_ -ge 0 -and $_ -le 10 } { return "<div class='container'><center><span class='toner-low'>$Level%</span></center></div>" }
        default { return $null }
    }
}

function Format-Status {
    param([string]$TcpStatus)
    
    $class = switch ($TcpStatus) {
        "Online"  { "online" }
        "Offline" { "offline" }
        default   { "error" }
    }
    return "<center><span class='$class'>$TcpStatus</span></center>"
}

function Send-Mail {
    param(
        [string]$MailFrom = "sender@example.com",
        [string[]]$MailTo = "recipient@example.com",
        [string]$Subject = "Report",
        [string]$Body = "Report", # If not set HTML Body
        [bool]$IsBodyHtml = $false,
        [string[]]$Attachments,
        [string[]]$CC,  # Carbon copy
        [string[]]$BCC, # Blind carbon copy
        [string]$SmtpServer = "smtp.example.com",
        [int]$SmtpPort = 25,
        [bool]$EnableSsl = $false,
        [bool]$UseDefaultCredentials = $true
    )

    $mail = $null
    $smtpClient = $null

    try {
        Write-Host "Trying to send a letter via $SmtpServer`:$SmtpPort..." -ForegroundColor Yellow
        
        $mail = [System.Net.Mail.MailMessage]::new()
        $mail.From = $MailFrom
        
        # Adding recipients
        foreach ($recipient in $MailTo) { $mail.To.Add($recipient) }
        
        # Adding copies
        if ($CC) { foreach ($ccRecipient in $CC) { $mail.CC.Add($ccRecipient) } }
        
        # Adding hidden copies
        if ($BCC) { foreach ($bccRecipient in $BCC) { $mail.Bcc.Add($bccRecipient) } }
        
        $mail.Subject = $Subject
        $mail.Body = $Body
        $mail.IsBodyHtml = $IsBodyHtml

        # Adding attachments
        if ($Attachments) {
            foreach ($attachmentPath in $Attachments) {
                if ($null -eq $attachmentPath -or $attachmentPath -like "") { continue }

                if (Test-Path $attachmentPath) {
                    $attachment = [System.Net.Mail.Attachment]::new($attachmentPath)
                    $mail.Attachments.Add($attachment)
                    Write-Host "Attachment added: $attachmentPath" -ForegroundColor Gray
                } else { Write-Warning "File not found: $attachmentPath" }
            }
        }

        # Setting up an SMTP client
        $smtpClient = [System.Net.Mail.SmtpClient]::new($SmtpServer, $SmtpPort)
        $smtpClient.EnableSsl = $EnableSsl
        $smtpClient.UseDefaultCredentials = $UseDefaultCredentials
        
        # Timeout
        $smtpClient.Timeout = $SmtpTimeoutMs # Timeout Smtp - Milliseconds (1 sec. = 1000 ms)

        # Send email
        $smtpClient.Send($mail)
        
        Write-Host "The letter has been sent successfully!" -ForegroundColor Green
        Write-Host "From: $MailFrom" -ForegroundColor Gray
        Write-Host "To: $($MailTo -join ', ')" -ForegroundColor Gray
        Write-Host "Subject: $Subject" -ForegroundColor Gray
        
        return
    }
    catch {
        Write-Error "Mail - Error sending email: $($_.Exception.Message)"
        if ($_.Exception.InnerException) { Write-Error "Mail - Internal error: $($_.Exception.InnerException.Message)" }
        return
    }
    finally {
        if ($mail) { 
            $mail.Dispose() 
            Write-Host "Letter resources freed" -ForegroundColor DarkGray
        }
        if ($smtpClient) { 
            $smtpClient.Dispose() 
            Write-Host "SMTP client resources have been released" -ForegroundColor DarkGray
        }
    }
}
#endregion

#region Initialization
try {
    Write-Log "=== Starting Printer Monitoring Version: $Version ==="

    # Load SharpSNMPLib.dll
    if ([System.Reflection.Assembly]::LoadFrom($DllPath)) { Write-Log "=== SharpSnmpLib: OK ===" }
    else { Write-Log "Failed to load Lextm.SharpSnmpLib" -Level "ERROR"
		throw "Missing Lextm.SharpSnmpLib Assembly; is it installed?" | Out-Null
    }

    # Load printer OID
    if (Test-Path $PrinterOIDPath) { $oidMapping = Import-PowerShellDataFile $PrinterOIDPath
        Write-Log "=== OID mapping: OK ==="
    } else { Write-Log "Printer OID file not found: $PrinterOIDPath" -Level "ERROR"
        throw "Printer OID file not found: $PrinterOIDPath" | Out-Null
    }

    # Interface Mode
    switch ($InterfaceMode) {
        "Console" {
            if (Test-Path $ConsoleFile) { $selectedFile = $ConsoleFile
                Write-Log "=== Printer list: OK ==="
            } else { Write-Log "Printer list file not load" -Level "ERROR"
                throw "Printer list file not load" | Out-Null
            }
        }
        "FullGui" {
            $PrinterRange = @()
            if (Test-Path $ScriptGuiXaml) { . $ScriptGuiXaml
                Write-Log "=== GUI Xaml: OK ==="
            } else { throw "GUI Xaml not load" | Out-Null }

            if ($selectedFile = Show-UserGUIXaml -Directory $PrinterListPath) { Write-Log "=== Printer list: OK ===" }
                else { Write-Log "Printer list file not load" -Level "ERROR"
                throw "Printer list file not load" | Out-Null
            }
        }
        "LightGui" {
            if ($selectedFile = Open-File $PrinterListPath) { Write-Log "=== Printer list: OK ===" }
                else { Write-Log "Printer list file not load" -Level "ERROR"
                throw "Printer list file not load" | Out-Null
            }
        }
    }

    # Import from (range/txt/csv)
    if ($PrinterRange -notlike $null) { $printers = $PrinterRange }
    elseif ($printers = Import-Csv -Path $selectedFile -Header Value) { Write-Log "=== Import Printers IP: OK ===" }
    else { Write-Log "=== Import Printers IP: Error ===" -Level "ERROR"
        throw "Import Printers IP: Error" | Out-Null
    }

    $totalPrinters = $printers.Value.Count
    
    if ($totalPrinters -eq 0) { Write-Log "No printers found in the list" -Level "ERROR"
        throw "No printers found in the list" | Out-Null
    }

    Write-Log "=== TCP test connection ==="
    $CheckPrinters = Test-TcpConnectionParallel -Devices $printers.Value -Port $TCPPort -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs -TimeoutMsTCP $TimeoutMsTCP -Threads $TCPThreads

    Write-Log "=== Monitoring $totalPrinters printers ==="

    $DataHtmlReport = [System.Collections.Generic.List[PSObject]]::new()
    $DataCsvReport = [System.Collections.Generic.List[PSObject]]::new($CsvBufferSize)
    if($CsvFileReport) {
        if ($PrinterRange.Value.Count -gt 0) {
            $filenamecsv = "$ReportDir\Printers_report_$($PrinterRange[0].Value)-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        }
        else {
            $filePrefixcsv = Split-Path -Path $selectedFile -Leaf
            $filePrefixcsv = $filePrefixcsv -replace '.txt', '' -replace '.csv', ''
            $filenamecsv = "$ReportDir\Printers_report_$filePrefixcsv-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        }
    }
#endregion

#region Main Processing
    $start = Get-Date
    $currentPrinter = 0

    ForEach ($currentPrinterIP in $CheckPrinters) {
        $currentPrinter++
        $printerIP = $currentPrinterIP.Device
        $TcpStatus = $null
        $oidSet = $null
        $pdisplay = $null
        $pstatus = $null
        $PrinterData = $null
        $tonerLevels = $null

        Write-Log "=== Checking printer: $currentPrinter of $totalPrinters ==="
        # Check printer availability(Ping)
        if ($currentPrinterIP.Connected -like 'false') {
            $TcpStatus = "Offline"
            Write-Host "$currentPrinter : $printerIP - Offline" -ForegroundColor Red
            Write-Log "Checking printer: $printerIP - Offline"
        }
        else {
            try {
                $TcpStatus = "Online"
                $PrinterData = Get-PrinterData -Target $printerIP -TimeoutMsUDP $TimeoutMsUDP -OidMapping $OidMapping -CurrentPrinter $CurrentPrinter
                $tonerLevels = Update-TonerLevels -PrinterData $PrinterData

                # Get printer status values
                if ($oidSet.Display) { $pdisplay = Get-SnmpBulkWalkWithEncoding -Target $printerIP -Oid $oidSet.Display -TimeoutMsUDP $TimeoutMsUDP }
                if ($oidSet.Status) { $pstatus = Get-SnmpBulkWalkWithEncoding -Target $printerIP -Oid $oidSet.Status -TimeoutMsUDP $TimeoutMsUDP }
            }
            catch {
                $TcpStatus = "Error"
                Write-Log "Error querying $printerIP : $_" -Level "ERROR"
            }
        }

        if($HtmlFileReport) {
            $htmlDisplay = foreach ($element in $pdisplay) { "<li>$element</li>" | Out-String -Stream }
            $htmlOutputDisplay = $htmlDisplay -join ""

            $htmlStatus = foreach ($element in $pstatus) { "<li>$element</li>" | Out-String -Stream }
            $htmlOutputStatus = $htmlStatus -join ""

            # Add information to HTML report
            $null = $DataHtmlReport.Add([PSCustomObject]@{
                "<span>IP</span>" = "<a class='printer-link' href='http://$printerIP' target='_blank'>$printerIP</a>"
                "<span>Name</span>" = "<a class='printer-link' href='http://$printerIP' target='_blank'>$($PrinterData.PName)</a>"
                "<span>Ping</span>" = Format-Status -TcpStatus $TcpStatus
                "<span>Model</span>" = Format-Value -Value $PrinterData.Model
                "<span>S/N</span>" = Format-Value -Value $PrinterData.Serial
                "<span>Black</span>" = Format-Value -Value $PrinterData.BlackCount
                "<span>Color</span>" = Format-Value -Value $PrinterData.ColorCount
                "<span>Total</span>" = Format-Value -Value $PrinterData.TotalCount
                "<span style='color:#00FFFF'>C</span> Toner"  = Format-TonerLevel -Level $tonerLevels.TC
                "<span style='color:#FD3DB5'>M</span> Toner"  = Format-TonerLevel -Level $tonerLevels.TM
                "<span style='color:#FFDE21'>Y</span> Toner"  = Format-TonerLevel -Level $tonerLevels.TY
                "<span style='color:#000000'>K</span> Toner"  = Format-TonerLevel -Level $tonerLevels.TK
                "<span style='color:#000000'>K</span> Drum"   = Format-TonerLevel -Level $tonerLevels.DKU
                "<span style='color:#00FFFF'>C</span><span style='color:#FD3DB5'>M</span><span style='color:#FFDE21'>Y</span><span style='color:#000000'>K</span> DrumKit" = Format-TonerLevel -Level $tonerLevels.DK
                "<span style='color:#FFDE21'>Display</span>" = if ($pdisplay.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputDisplay</ul></div>" } else { "" }
                "<span style='color:#FFDE21'>Active Alerts</span>" = if ($pstatus.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputStatus</ul></div>" } else { "" }
                # Add other columns similarly
            })
        }

        if($CsvFileReport) {
            # Add information to CSV report
            $null = $DataCsvReport.Add([PSCustomObject]@{
                "IP"    = "$printerIP"
                "Name"  = $PrinterData.PName
                "Ping"  = "$TcpStatus"
                "Model" = $PrinterData.Model
                "S/N"   = $PrinterData.Serial
                "Black" = $PrinterData.BlackCount
                "Color" = $PrinterData.ColorCount
                "Total" = $PrinterData.TotalCount
                "C Toner %" = $tonerLevels.TC
                "M Toner %" = $tonerLevels.TM
                "Y Toner %" = $tonerLevels.TY
                "K Toner %" = $tonerLevels.TK
                "K Drum %"  = $tonerLevels.DKU
                "CMYK DrumKit %" = $tonerLevels.DK
                #"Display" = if ($pdisplay.Length -ne 0) { "$pdisplay" } else { "" }        #It is necessary to change the data output format in the SNMP BulkMWalk function
                #"Active Alerts" = if ($pstatus.Length -ne 0) { "$pstatus" } else { "" }    #It is necessary to change the data output format in the SNMP BulkMWalk function
                # Add other columns similarly
            })

            if ($DataCsvReport.Count -ge $CsvBufferSize) {
                $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
                $DataCsvReport.Clear()
            }
        }
    }

    if ($DataCsvReport.Count -gt 0) {
        $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
        $DataCsvReport.Clear()
    }

    $end = Get-Date
    $duration = $end - $start
    Write-Log "=== Monitoring completed in $($duration.TotalSeconds) seconds ==="
#endregion

#region Report Generation
    # HTML
    If($HtmlFileReport) {
        if (Test-Path $HtmlHeader) { . $HtmlHeader
            Write-Log "=== Html Header: OK ==="
        } else { throw "Html Header not load" | Out-Null }

        $htmlContent = $DataHtmlReport | ConvertTo-Html -Title "Printer Status Report" -Head $Header -PostContent "<p>Report generated: $(Get-Date)</p>"
        # Fix HTML encoding
        $htmlContent = $htmlContent -replace '&lt;', '<' -replace '&#39;', "'" -replace '&gt;', '>'

        if ($PrinterRange.Value.Count -gt 0) {
            $filenamehtml = "$ReportDir\Printers_report_$($PrinterRange[0].Value)-$(Get-Date -Format 'yyyyMMddHHmmss').html"
            $htmlContent | Set-Content -Path $filenamehtml -Force -Encoding UTF8
        }
        else {  
            $filePrefixhtml = Split-Path -Path $selectedFile -Leaf
            $filePrefixhtml = $filePrefixhtml -replace '.txt', '' -replace '.csv', ''
            $filenamehtml = "$ReportDir\Printers_report_$filePrefixhtml-$(Get-Date -Format 'yyyyMMddHHmmss').html"
            $htmlContent | Set-Content -Path $filenamehtml -Force -Encoding UTF8
        }

        if (Test-Path $filenamehtml) {
            $DataHtmlReport.Clear()
            Write-Log "Report generated: $filenamehtml"
            if (-not $MailSend) { Start-Process $filenamehtml }
        }
        else { Write-Log "Failed to generate HTML report" -Level "ERROR" }
    }
    
    # CSV
    if($CsvFileReport) {
        if (Test-Path $filenamecsv) {
            Write-Log "Report generated: $filenamecsv"
            if (-not $MailSend) { Start-Process $filenamecsv }
        }
        else { Write-Log "Failed to generate CSV report" -Level "ERROR" }
    }
    
    # Mail
    if ($MailSend) {
        Send-Mail -MailFrom $MailFrom -MailTo @($MailTo) -Subject $Subject -Attachments @($filenamehtml, $filenamecsv) -Body $MailHtmlBody -IsBodyHtml $true -SmtpServer $SmtpServer -SmtpPort $SmtpPort
    }
}
catch {
    Write-Log "Script failed: $_" -Level "ERROR"
    exit 1
}
finally {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
#endregion