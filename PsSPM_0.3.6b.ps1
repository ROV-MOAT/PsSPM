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
    Date: 2026-05-25
    
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

[string]$Version = "PsSPM 0.3.6b-20260525" #Y/M/D
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
[int]$SnmpMaxAttempts = 3   # Retry Count
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
[string]$HtmlPath = "$PSScriptRoot\lib\PsSPM_html.psm1"
[string]$ScriptGuiXaml = "$PSScriptRoot\lib\PsSPM_wpf.psm1"
[string]$SnmpPath = "$PSScriptRoot\lib\PsSPM_snmp.psm1"
[string]$MailPath = "$PSScriptRoot\lib\PsSPM_mail.psm1"
[string]$CollectorPath = "$PSScriptRoot\lib\PsSPM_collector.psm1"
[string]$ReportPath = "$PSScriptRoot\lib\PsSPM_report.psm1"
[string]$TcpPath = "$PSScriptRoot\lib\PsSPM_tcp.psm1"
[string]$PrinterListPath = "$PSScriptRoot\ip"
[string]$ReportDir ="$PSScriptRoot\report"
[string]$DllPath = "$PSScriptRoot\lib\SharpSnmpLib.dll"
[string]$selectedFile = $null
[string]$filenamecsv = $null
[string]$filenamehtml = $null
[array]$PrinterRange = @()
[array]$printerIPs = @()

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
if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = '{0} [{1}] - {2}' -f $timestamp, $Level, $Message
    if ($WriteLog) {
        Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8  
    }
    if ($ShowLog) {
        switch ($Level) {
            "ERROR"    { Write-Host $logMessage -ForegroundColor Red }
            "WARNING"  { Write-Host $logMessage -ForegroundColor Yellow }
            #"BULK"     { Write-Host $logMessage -ForegroundColor Blue }
            default    { Write-Host $logMessage -ForegroundColor Green }
        }
    }
}

function Test-AndLoad {
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [Parameter(Mandatory)]
        [string] $Name,
        [ValidateSet("Assembly", "OidMap", "Module")]
        [string] $LoadType = "Module"
    )

    if (-not (Test-Path $Path)) {
        Write-Log "$Name not found: $Path" -Level "ERROR"
        throw "$Name not found: $Path"
    }

    try {
        switch ($LoadType) {
            "Assembly" {
                [System.Reflection.Assembly]::LoadFrom($Path) | Out-Null
                Write-Log "$Name : OK"
                return
            }
            "OidMap" {
                $data = Import-PowerShellDataFile -Path $Path
                Write-Log "$Name : OK"
                return $data
            }
            "Module" {

                if (-not (Get-Module -Name (Split-Path $Path -Leaf) -ErrorAction SilentlyContinue)) {
                    $module = Import-Module $Path -Force
                    Write-Log "$Name : OK"
                }
                return $module
            }
        }
    }
    catch {
        Write-Log "Failed to load $Name — $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Get-ReportFilename {
    param(
        [string]$Ext
    )

    $stamp = Get-Date -Format 'yyyyMMddHHmmss'

    if ($PrinterRange.Count -gt 0) {
        $firstIP = if ($PrinterRange[0] -is [PSCustomObject]) {
            $PrinterRange[0].Value
        } else {
            $PrinterRange[0]
        }
        return "$ReportDir\PsSPM_report_${firstIP}_$stamp.$Ext"
    }

    $prefix = [IO.Path]::GetFileNameWithoutExtension($selectedFile)
    return "$ReportDir\PsSPM_report_${prefix}_$stamp.$Ext"
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
#endregion Helper Functions

#region Initialization
Write-Log "=== Starting Printer Monitoring Version: $Version ==="

Test-AndLoad -Path $DllPath -Name "SharpSnmpLib" -LoadType "Assembly"
Test-AndLoad -Path $TcpPath -Name "TCP function" -LoadType "Module"
Test-AndLoad -Path $SnmpPath -Name "SNMP function" -LoadType "Module"
Test-AndLoad -Path $CollectorPath -Name "Collector function" -LoadType "Module"
Test-AndLoad -Path $ReportPath -Name "Report function" -LoadType "Module"

$oidMapping = Test-AndLoad -Path $PrinterOIDPath -Name "OID Mapping" -LoadType "OidMap"

if ($MailSend) { Test-AndLoad -Path $MailPath -Name "Mail Function" -LoadType "Module" }

try {
    #region Interface mode & source selection
    switch ($InterfaceMode) {
        'Console' {
            if (-not (Test-Path -LiteralPath $ConsoleFile)) {
                Write-Log "Printer list file not found: $ConsoleFile" -Level "ERROR"
                throw "Printer list file not found"
            }

            $selectedFile = $ConsoleFile
            Write-Log "Printer list: OK (Console mode)"
        }
        'FullGui' {
            Test-AndLoad -Path $ScriptGuiXaml -Name "GUI function" -LoadType "Module"

            $selectedFile = Show-UserGUIXaml -Directory $PrinterListPath
            if (-not $selectedFile) {
                Write-Log "Printer list file not selected in GUI" -Level "ERROR"
                throw "Printer list file not selected"
            }

            if (-not (Test-Path -LiteralPath $selectedFile)) {
                Write-Log "Selected printer list file not found: $selectedFile" -Level "ERROR"
                throw "Selected printer list file not found"
            }

            Write-Log "Printer list: OK (FullGui mode)"
        }
        default {
            Write-Log "Unsupported InterfaceMode: $InterfaceMode" -Level "ERROR"
            throw "Unsupported InterfaceMode: $InterfaceMode"
        }
    }
    #endregion Interface mode & source selection

    #region Import printer IPs (range / txt / csv)
    if ($PrinterRange -and $PrinterRange.Count -gt 0) {
        if ($PrinterRange[0] -is [PSCustomObject] -and ($PrinterRange[0] | Get-Member -Name Value -ErrorAction SilentlyContinue)) {
            $printerIPs = $PrinterRange.Value
        }
        else {
            $printerIPs = $PrinterRange
        }

        $printerIPs = $printerIPs | Where-Object { $_ -and $_.ToString().Trim() -ne '' } | Select-Object -Unique
        Write-Log "=== Using printer range: $($printerIPs.Count) IPs ==="
    }
    elseif ($selectedFile) {
        $ext = [IO.Path]::GetExtension($selectedFile).ToLower()
        
        switch ($ext) {
            '.csv' {
                try {
                    $importedData = Import-Csv -Path $selectedFile -ErrorAction Stop

                    if ($importedData.Count -eq 0) { throw "CSV file is empty" }

                    if ($importedData[0].PSObject.Properties.Name -contains 'Value') {
                        $printerIPs = $importedData.Value
                    } else {
                        $firstProp = $importedData[0].PSObject.Properties.Name | Select-Object -First 1
                        $printerIPs = $importedData.$firstProp
                    }

                    Write-Log "=== Import Printers IP (CSV): OK ($($printerIPs.Count) IPs) ==="
                }
                catch {
                    Write-Log "CSV import error: $_" -Level "ERROR"
                    throw
                }
            }
            '.txt' {
                $printerIPs = Get-Content -Path $selectedFile -Encoding UTF8 | Where-Object { $_ -and $_.Trim() -ne '' }
                Write-Log "=== Import Printers IP (TXT): OK ($($printerIPs.Count) IPs) ==="
            }
            default {
                throw "Unsupported file type: $ext"
            }
        }
    }
    else {
        Write-Log "=== No printer source found (no range and no file) ===" -Level "ERROR"
        throw "No printer source found"
    }

    $totalPrinters = $printerIPs.Count
    if ($totalPrinters -le 0) {
        Write-Log "No printers found in the list after import" -Level "ERROR"
        throw "No printers found in the list"
    }
    #endregion Import printer IPs

    #region TCP check
    $tcpstart = Get-Date
    Write-Log "=== TCP test connection (Threads: $TCPThreads, Port: $TCPPort) ==="

    $CheckPrinters = Test-TcpConnectionParallel -Devices $printerIPs -Port $TCPPort -TcpTimeoutMs $TcpTimeoutMs -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs -Threads $TCPThreads
    #$CheckPrinters
    if (-not $CheckPrinters) {
        Write-Log "TCP check returned no results" -Level "ERROR"
        throw "TCP check returned no results"
    }    
    $tcpend = Get-Date
    $tcpduration = $tcpend - $tcpstart
    Write-Log "=== TCP check completed in $([math]::Round($tcpduration.TotalSeconds,2)) seconds ==="

    #endregion TCP check

    #region Initialize report collections
    $DataHtmlReport = [System.Collections.Generic.List[string]]::new()
    $DataCsvReport  = [System.Collections.Generic.List[PSObject]]::new($CsvBufferSize)

    if ($CsvFileReport) { $filenamecsv = Get-ReportFilename -Prefix "report" -Ext "csv" }
    #endregion Initialize report collections

    #region Main Processing (SNMP)
    $snmpstart = Get-Date
    $totalPrinters = $CheckPrinters.Count

    Write-Log "Starting parallel SNMP check of $totalPrinters printers with Threads: $SNMPThreads"

    $SnmpResults = Get-PrintersDataParallel -CheckPrinters $CheckPrinters -OidMapping $OidMapping -ModelPatterns $modelPatterns -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs -MaxThreads $SNMPThreads

    if (-not $SnmpResults) {
        Write-Log "Get-PrintersDataParallel returned no results" -Level "ERROR"
        throw "SNMP data collection returned no results"
    }

    $SortAllResults = $SnmpResults | Sort-Object { Convert-IPToUInt32 $_.IPAddress }

    foreach ($result in $SortAllResults) {
        #$result.PrinterData.GetType()
        if ($HtmlFileReport) {
            Add-HtmlPrinterRowString -DataHtmlReport $DataHtmlReport -Collector $result
        }

        if ($CsvFileReport) {
            Add-CsvPrinterRow -DataCsvReport $DataCsvReport -PrinterIP $result.IPAddress -TcpStatus $result.TCPStatus -PrinterData $result.PrinterData `
                            -TonerLevels $result.TonerLevels -PrinterErrors $result.SnmpErrors -CsvPath $filenamecsv -CsvBufferSize $CsvBufferSize
        }

    }

    if ($CsvFileReport -and $filenamecsv) {
        if ($DataCsvReport.Count -gt 0) {
            $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
            $DataCsvReport.Clear()
        }
    }

    $snmpend = Get-Date
    $snmpduration = $snmpend - $snmpstart
    Write-Log "=== SNMP check completed in $([math]::Round($snmpduration.TotalSeconds,2)) seconds ==="
    #endregion Main processing (SNMP)

    #region Report Generation
    # HTML
    If($HtmlFileReport) {
        Test-AndLoad -Path $HtmlPath -Name "HTML function" -LoadType "Module"

        $filenamehtml = Get-ReportFilename -Prefix "report" -Ext "html"
        $FinalHtml | Set-Content -Path $filenamehtml -Force -Encoding UTF8

        if (Test-Path -LiteralPath $filenamehtml) {
            $DataHtmlReport.Clear()
            Write-Log "HTML report generated: $filenamehtml"

            if (-not $MailSend) { Start-Process -FilePath $filenamehtml }
        }
        else {
            Write-Log "Failed to generate HTML report: $filenamehtml" -Level "ERROR"
        }
    }
    
    # CSV
    if ($CsvFileReport -and $filenamecsv) {
        if (Test-Path -LiteralPath $filenamecsv) {
            Write-Log "CSV report generated: $filenamecsv"

            if (-not $MailSend) { Start-Process -FilePath $filenamecsv }
        }
        else {
            Write-Log "Failed to generate CSV report: $filenamecsv" -Level "ERROR"
        }
    }
    
    # Mail
    if ($MailSend) {
        $attachments = @()

        if ($filenamehtml -and (Test-Path -LiteralPath $filenamehtml)) { $attachments += $filenamehtml }
        if ($filenamecsv  -and (Test-Path -LiteralPath $filenamecsv )) { $attachments += $filenamecsv  }

        Send-UniversalMail -MailFrom $MailFrom -MailTo @($MailTo) -CC @($CC) -BCC @($BCC) -Subject $Subject -Attachments $attachments -UseDefaultCredentials `
            -Username $MailUser -Password $MailPass -Body $MailHtmlBody -IsBodyHtml $true -SmtpServer $SmtpServer -SmtpPort $SmtpPort -EnableSsl $EnableSsl -SmtpTimeoutMs $SmtpTimeoutMs
    }
    #endregion Report generation
}
catch {
    Write-Log "Script failed: $_" -Level "ERROR"
    exit 1
}
finally {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
#endregion Initialization