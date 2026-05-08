<#
.SYNOPSIS
    PowerShell SNMP Printer Monitoring and Reporting Script / PsSPM (ROV-MOAT)
    C# SNMP Library is used - https://github.com/lextudio/sharpsnmplib

.DESCRIPTION
    Checks printer status via SNMP and generates CSV, HTML report with toner levels, counters, and information

.LICENSE
    Distributed under the MIT License. See the accompanying LICENSE file or https://github.com/ROV-MOAT/PsSPM/blob/main/LICENSE
    
.NOTES
    Version: 0.3.6b
    Author: Oleg Ryabinin + AI
    Date: 2026-05-08
    
    MESSAGE:
    Powershell 5+

    CHANGELOG:
    Ver. 0.3.6b
    + RunspacePool (Get-PrintersDataParallel)
    Optimization

    Ver. 0.3.5b
    + Export from HTML
    + Copy text - MAC, S/N
    + Search in / Filter
    + MAC address
    + Mail
    + Interface Mode (Console, FullGui)
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
    [ValidateSet("Console", "FullGui")]
    [string]$InterfaceMode = "FullGui",
    [string]$ConsoleFile = "",
    [bool]$MailSend = $false,
    [string]$MailUser = "",
    [string]$MailPass = ""
)

[string]$Version = "PsSPM 0.3.6b-20260508" #Y/M/D
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# REPORT
[bool]$HtmlFileReport = $true   # On/Off HTML Report out file
[bool]$CsvFileReport  = $false  # On/Off CSV Report out file
[int]$CsvBufferSize = 50        # Number of lines in buffer to write to CSV file
[bool]$Script:GetBulkWalk = $false      #Get Display/Active Alerts

# LOG
[bool]$WriteLog = $false   # On/Off Logging Function
[bool]$ShowLog = $true   # On/Off console output
[string]$LogDir = "$PSScriptRoot\log"
[string]$LogFile = "$logDir\PsSPM_log_$(Get-Date -Format 'yyyyMMddHHmmss').log"

# SNMP
[int]$SnmpTimeoutMs = 5000  # Timeout SNMP - Milliseconds (1 sec. = 1000 ms) / 0 = infinity
[int]$SnmpMaxAttempts = 4   # Retry Count
[int]$SnmpDelayMs = 2000    # Milliseconds (1 sec. = 1000 ms)
[int]$SNMPThreads = 50      # Maximum number of running SNMP threads

# TCP
[int]$TCPPort = 80         # TCP port for check link
[int]$TcpTimeoutMs = 1000   # Timeout TcpClient - Milliseconds (1 sec. = 1000 ms)
[int]$MaxRetries   = 4     # Retry Count
[int]$RetryDelayMs = 2000  # Milliseconds (1 sec. = 1000 ms)
[int]$TCPThreads = 50      # Maximum number of running TCP threads

# Mail
[string]$MailFrom = "example@example.com"
[string[]]$MailTo = "example@example.com"
[string]$Subject = "PsSPM Report"
[string]$SmtpServer = "mail.example.com"
[string[]]$CC = ""
[string[]]$BCC = ""
[int]$SmtpPort = 25
[bool]$EnableSsl = $false
[int]$SmtpTimeoutMs = 10000

# Path
[string]$PrinterOIDPath = "$PSScriptRoot\lib\PsSPM_oid.psd1"
[string]$ScriptGuiXaml = "$PSScriptRoot\lib\PsSPM_wpf.ps1"
[string]$HtmlHeader = "$PSScriptRoot\lib\PsSPM_html.ps1"
[string]$MailPath = "$PSScriptRoot\lib\PsSPM_mail.ps1"
[string]$SnmpPath = "$PSScriptRoot\lib\PsSPM_snmp.ps1"
[string]$CollectorPath = "$PSScriptRoot\lib\PsSPM_collector.ps1"
[string]$TcpPath = "$PSScriptRoot\lib\PsSPM_tcp.ps1"
[string]$PrinterListPath = "$PSScriptRoot\ip"
[string]$ReportDir ="$PSScriptRoot\report"
[string]$DllPath = "$PSScriptRoot\lib\SharpSnmpLib.dll"
[string]$selectedFile = $null
[string]$filenamecsv = $null
[string]$filenamehtml = $null

# Printer
[hashtable]$modelPatterns = @{
    "333" = "333"
    "B60" = "B60"
    "B61" = "B60"
    "B70" = "B60"
    "M40" = "M40"
    "B80" = "B80"
    "365" = "365"
    "594" = "365"
    "4622" = "365"
    "C40" = "C40"
    "650" = "660"
    "660" = "660"
    "651" = "651"
    "T79" = "T79"
    "332" = "332"
    "670" = "670"
    "450" = "450"
    "C60" = "C60"
    "533" = "B60"
    "MX622" = "622"
    "611" = "611"
    "M42" = "M42"
    "M25" = "M25"
    "T73" = "T73"
    "350" = "350"
    "B40" = "B60"
    "M28" = "M25"
}
#region Helper Functions
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir | Out-Null }

# Logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] - $Message"
    
    if ($WriteLog) {
        Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8  
    }
    if ($ShowLog) {
        switch ($Level) {
            "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
            "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
            default   { Write-Host $logMessage -ForegroundColor Green }
        }
    }
}

function Convert-IPToUInt32 {
    param([string]$IPAddress)

    if ([string]::IsNullOrEmpty($IPAddress)) { return [uint64]0 }

    try {
        $bytes = [System.Net.IPAddress]::Parse($IPAddress).GetAddressBytes()
        if ($bytes.Count -eq 4) {
            return [uint64]$bytes[0] * 256 * 256 * 256 +
                   [uint64]$bytes[1] * 256 * 256 +
                   [uint64]$bytes[2] * 256 +
                   [uint64]$bytes[3]
        }
    }
    catch { return [uint64]0 }
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

function Format-TonerLevel {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Level
    )

    if ($null -eq $Level -or $Level -like "") { return $null }
    if ($Level -like 'Error') { return "<span class='error'>Error</span>" }

    try { $tonerValue = [int]$Level } catch { return $null }

    switch ($tonerValue) {
        { $_ -gt 49 } { return "<span class='toner-high'>$Level%</span>" }
        { $_ -gt 10 -and $_ -le 49 } { return "<span class='toner-medium'>$Level%</span>" }
        { $_ -ge 0 -and $_ -le 10 } { return "<span class='toner-low'>$Level%</span>" }
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
#endregion

#region Initialization
try {
    Write-Log "=== Starting Printer Monitoring Version: $Version ==="

    # Load SharpSNMPLib.dll
    if ([System.Reflection.Assembly]::LoadFrom($DllPath)) { Write-Log "=== SharpSnmpLib: OK ===" }
    else { Write-Log "Failed to load Lextm.SharpSnmpLib" -Level "ERROR"
		throw "Missing Lextm.SharpSnmpLib Assembly; is it installed?"
    }

    # Load printer OID
    if (Test-Path $PrinterOIDPath) { $oidMapping = Import-PowerShellDataFile $PrinterOIDPath
        Write-Log "=== OID mapping: OK ==="
    } else { Write-Log "Printer OID file not found: $PrinterOIDPath" -Level "ERROR"
        throw "Printer OID file not found: $PrinterOIDPath"
    }

    # Load Mail
    if ($MailSend) {
        if (Test-Path $MailPath) { . $MailPath
            Write-Log "=== Mail Function: OK ==="
        } else { Write-Log "Mail function file not found: $MailPath" -Level "ERROR"
            throw "Mail function file not found: $MailPath"
        }
    }
    
    # SNMP
    if (Test-Path $SnmpPath) { 
        . $SnmpPath
        Write-Log "=== SNMP Function: OK ==="
    } else { 
        Write-Log "SNMP function file not found: $SnmpPath" -Level "ERROR"
        throw "SNMP function file not found: $SnmpPath"
    }

    if (Test-Path $CollectorPath) { 
        . $CollectorPath
        Write-Log "=== Collector Function: OK ==="
    } else { 
        Write-Log "Collector function file not found: $CollectorPath" -Level "ERROR"
        throw "Collector function file not found: $CollectorPath"
    }

    # Interface Mode
    $PrinterRange = @()

    switch ($InterfaceMode) {
        "Console" {
            if (Test-Path $ConsoleFile) { 
                $selectedFile = $ConsoleFile
                Write-Log "=== Printer list: OK ==="
            } else { 
                Write-Log "Printer list file not load" -Level "ERROR"
                throw "Printer list file not load"
            }
        }
        "FullGui" {
            if (Test-Path $ScriptGuiXaml) { 
                . $ScriptGuiXaml
                Write-Log "=== WPF GUI: OK ==="
            } else { 
                throw "WPF GUI not load"
            }

            if ($selectedFile = Show-UserGUIXaml -Directory $PrinterListPath) { 
                Write-Log "=== Printer list: OK ===" 
            } else { 
                Write-Log "Printer list file not load" -Level "ERROR"
                throw "Printer list file not load"
            }
        }
    }

    # Import from (range/txt/csv)
    $printerIPs = @()

    if ($PrinterRange -and $PrinterRange.Count -gt 0) { 
        if ($PrinterRange[0] -is [PSCustomObject] -and ($PrinterRange[0] | Get-Member -Name Value -ErrorAction SilentlyContinue)) {
            $printerIPs = $PrinterRange.Value
        } else {
            $printerIPs = $PrinterRange
        }
        Write-Log "=== Using printer range: $($printerIPs.Count) IPs ==="
    }
    elseif ($selectedFile -and (Test-Path $selectedFile)) {
        try {
            $importedData = Import-Csv -Path $selectedFile -Header Value
            if ($importedData -and $importedData.Count -gt 0) {
                $printerIPs = $importedData.Value
                Write-Log "=== Import Printers IP: OK ($($printerIPs.Count) IPs) ==="
            } else {
                throw "CSV file is empty"
            }
        }
        catch {
            Write-Log "=== Import Printers IP: Error: $_ ===" -Level "ERROR"
            throw "Import Printers IP: Error"
        }
    }
    else {
        Write-Log "=== No printer source found ===" -Level "ERROR"
        throw "No printer source found"
    }

    $totalPrinters = $printerIPs.Count

    if ($totalPrinters -eq 0) { 
        Write-Log "No printers found in the list" -Level "ERROR"
        throw "No printers found in the list"
    }

    # Test connection
    if (Test-Path $TcpPath) { 
        . $TcpPath
        Write-Log "=== TCP Function: OK ==="
    } else { 
        Write-Log "TCP function file not found: $TcpPath" -Level "ERROR"
        throw "TCP function file not found: $TcpPath"
    }

    Write-Log "=== TCP test connection (Threads: $TCPThreads) ==="
    $CheckPrinters = Test-TcpConnectionParallel -Devices $printerIPs -Port $TCPPort -TcpTimeoutMs $TcpTimeoutMs -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs -Threads $TCPThreads

    Write-Log "=== Monitoring $totalPrinters printers ==="

    # Initialize report collections
    $DataHtmlReport = [System.Collections.Generic.List[PSObject]]::new()
    $DataCsvReport = [System.Collections.Generic.List[PSObject]]::new($CsvBufferSize)

    # Create CSV report file if needed
    if ($CsvFileReport) {
        try {
            if ($PrinterRange -and $PrinterRange.Count -gt 0) {
                $firstIP = if ($PrinterRange[0] -is [PSCustomObject] -and ($PrinterRange[0] | Get-Member -Name Value)) {
                    $PrinterRange[0].Value
                } else {
                    $PrinterRange[0]
                }
                $filenamecsv = "$ReportDir\PsSPM_report_$firstIP-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            }
            elseif ($selectedFile) {
                $filePrefixcsv = Split-Path -Path $selectedFile -Leaf
                $filePrefixcsv = $filePrefixcsv -replace '\.txt$|\.csv$', ''
                $filenamecsv = "$ReportDir\PsSPM_report_$filePrefixcsv-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            }
            else {
                $filenamecsv = "$ReportDir\PsSPM_report_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            }
            
            Write-Log "CSV report will be saved to: $filenamecsv"
            # NOTE: Export CSV data AFTER populating $DataCsvReport, not here
        }
        catch {
            Write-Log "Failed to create CSV report: $_" -Level "WARNING"
        }
    }
#endregion

#region Main Processing
$start = Get-Date
$totalPrinters = $CheckPrinters.Count

Write-Log "Starting parallel check of $totalPrinters printers with Threads: $SNMPThreads"

$allResults = Get-PrintersDataParallel -CheckPrinters $CheckPrinters -OidMapping $OidMapping -ModelPatterns $modelPatterns -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs -MaxThreads $SNMPThreads
$sortAllResults = $allResults | Sort-Object { Convert-IPToUInt32 $_.IPAddress }

    # Обработка результатов
    $currentPrinter = 0
    foreach ($result in $sortAllResults) {
        $currentPrinter++
        $printerIP = $result.IPAddress
        $TcpStatus = $result.TCPStatus
        $PrinterData = $result.PrinterData
        $tonerLevels = $result.TonerLevels
        $printerErrors = $result.AllErrors

        if($HtmlFileReport) {
            $htmlDisplay = foreach ($element in $PrinterData.Display) { "<li>$element</li>" | Out-String -Stream }
            $htmlOutputDisplay = $htmlDisplay -join ""

            $htmlStatus = foreach ($element in $PrinterData.Status) { "<li>$element</li>" | Out-String -Stream }
            $htmlOutputStatus = $htmlStatus -join ""

            $htmlError = foreach ($element in $printerErrors) { "<li>$element</li>" | Out-String -Stream }
            $htmlOutputError = $htmlError -join ""

            # Add information to HTML report
            $null = $DataHtmlReport.Add([PSCustomObject]@{
                "<span>IP</span>" = "<a class='printer-link' href='http://$printerIP' target='_blank'>$printerIP</a>"
                "<span>Ping</span>" = Format-Status -TcpStatus $TcpStatus
                "<span>Name</span>" = "<a class='printer-link' href='http://$printerIP' target='_blank'>$($PrinterData.PName)</a>"
                "<span>MAC</span>" = "<span style='cursor: copy;' onclick='copyText(this)'>$(Format-Value -Value $PrinterData.PMac)</span>"
                "<span>Model</span>" = Format-Value -Value $PrinterData.Model
                "<span>S/N</span>" = "<span style='cursor: copy;' onclick='copyText(this)'>$(Format-Value -Value $PrinterData.Serial)</span>"
                "<i class='fa-regular fa-file-lines me-1'></i><span>Black</span>" = Format-Value -Value $PrinterData.BlackCount
                "<i class='fa-regular fa-file-lines me-1'></i><span>Color</span>" = Format-Value -Value $PrinterData.ColorCount
                "<i class='fa-regular fa-file-lines me-1'></i><span>Total</span>" = Format-Value -Value $PrinterData.TotalCount
                "<span style='color:#00FFFF'>C</span> Toner"  = if (-not($tonerLevels.TC -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TC)</span>"} else { "" }
                "<span style='color:#FD3DB5'>M</span> Toner"  = if (-not($tonerLevels.TM -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TM)</span>"} else { "" }
                "<span style='color:#FFDE21'>Y</span> Toner"  = if (-not($tonerLevels.TY -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TY)</span>"} else { "" }
                "<span style='color:#000000'>K</span> Toner"  = if (-not($tonerLevels.TK -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TK)</span>"} else { "" }
                "<span style='color:#00FFFF'>C</span><span style='color:#FD3DB5'>M</span><span style='color:#FFDE21'>Y</span><span style='color:#000000'>K</span> DrumKit" = `
                    if (-not($tonerLevels.DC -like $null -and $tonerLevels.DM -like $null -and $tonerLevels.DY -like $null -and $tonerLevels.DK -like $null -and $tonerLevels.DKU -like $null)) {
                    "<span class='container'>$(Format-TonerLevel -Level $tonerLevels.DC) $(Format-TonerLevel -Level $tonerLevels.DM) $(Format-TonerLevel -Level $tonerLevels.DY) `
                    $(Format-TonerLevel -Level $tonerLevels.DKU) $(Format-TonerLevel -Level $tonerLevels.DK)</span>"} else { "" }
                "<span style='color:#FFDE21'>Display</span>" = if ($PrinterData.Display.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputDisplay</ul></div>" } else { "" }
                "<span style='color:#FFDE21'>Active Alerts</span>" = if ($PrinterData.Status.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputStatus</ul></div>" } else { "" }
                "<span style='color:#FFDE21'>E</span>" = if ($printerErrors.Length -ne 0) {"<div class='tooltip'><a class='show-link' href='' target='_blank'>E</a><ul class='tooltiptext'>$htmlOutputError</ul></div>" } else { "" }
                # Add other columns similarly
            })
        }

        if($CsvFileReport) {
            # Add information to CSV report
            $null = $DataCsvReport.Add([PSCustomObject]@{
                "IP"    = $printerIP
                "Ping"  = $TcpStatus
                "Name"  = $PrinterData.PName
                "MAC"   = $PrinterData.PMac
                "Model" = $PrinterData.Model
                "S/N"   = $PrinterData.Serial
                "Black" = $PrinterData.BlackCount
                "Color" = $PrinterData.ColorCount
                "Total" = $PrinterData.TotalCount
                "C Toner %" = $tonerLevels.TC
                "M Toner %" = $tonerLevels.TM
                "Y Toner %" = $tonerLevels.TY
                "K Toner %" = $tonerLevels.TK
                "CMYK DrumKit %" = "$($tonerLevels.DC) $($tonerLevels.DM) $($tonerLevels.DY) $($tonerLevels.DKU) $($tonerLevels.DK)"
                "Display" = if ($PrinterData.Display.Length -ne 0) { "$($PrinterData.Display)" } else { "" }
                "Active Alerts" = if ($PrinterData.Status.Length -ne 0) { "$($PrinterData.Status)" } else { "" }
                "Collector Error" = "$printerErrors"
                # Add other columns similarly
            })

            if ($DataCsvReport.Count -ge $CsvBufferSize) {
                $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
                $DataCsvReport.Clear()
            }
        }
    }

    if ($DataCsvReport.Count -gt 0 -and $filenamecsv) {
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
        } else { throw "Html Header not load" }

        $htmlContent = $DataHtmlReport | ConvertTo-Html -Title "Printer Status Report" -Head $Header -Body $ExBody -PostContent $ExBottom
        # Fix HTML encoding
        $htmlContent = $htmlContent -replace '&lt;', '<' -replace '&#39;', "'" -replace '&gt;', '>' -replace'<table>', '<table id="PrinterTable">'

        if ($PrinterRange -and $PrinterRange.Count -gt 0) {
            $firstIP = if ($PrinterRange[0] -is [PSCustomObject] -and ($PrinterRange[0] | Get-Member -Name Value)) {
                $PrinterRange[0].Value
            } else {
                $PrinterRange[0]
            }
            $filenamehtml = "$ReportDir\PsSPM_report_$firstIP-$(Get-Date -Format 'yyyyMMddHHmmss').html"
            $htmlContent | Set-Content -Path $filenamehtml -Force -Encoding UTF8
        }
        else {
            $filePrefixhtml = Split-Path -Path $selectedFile -Leaf
            $filePrefixhtml = $filePrefixhtml -replace '.txt', '' -replace '.csv', ''
            $filenamehtml = "$ReportDir\PsSPM_report_$filePrefixhtml-$(Get-Date -Format 'yyyyMMddHHmmss').html"
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
        Send-UniversalMail -MailFrom $MailFrom -MailTo @($MailTo) -CC @($CC) -BCC @($BCC) -Subject $Subject -Attachments @($filenamehtml, $filenamecsv) -UseDefaultCredentials `
            -Username $MailUser -Password $MailPass -Body $MailHtmlBody -IsBodyHtml $true -SmtpServer $SmtpServer -SmtpPort $SmtpPort -EnableSsl $EnableSsl -SmtpTimeoutMs $SmtpTimeoutMs
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