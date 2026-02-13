<#
.SYNOPSIS
    PowerShell SNMP Printer Monitoring and Reporting Script / PsSPM (ROV-MOAT)
    C# SNMP Library is used - https://github.com/lextudio/sharpsnmplib

.DESCRIPTION
    Checks printer status via SNMP and generates CSV, HTML report with toner levels, counters, and information

.LICENSE
    Distributed under the MIT License. See the accompanying LICENSE file or https://github.com/ROV-MOAT/PsSPM/blob/main/LICENSE
    
.NOTES
    Version: 0.3.5b
    Author: Oleg Ryabinin + AI
    
    MESSAGE:
    Powershell 5+

    CHANGELOG:

    Ver. 0.3.5b
	+ Copy text - MAC, S/N
	+ Search in
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

[string]$Version = "PsSPM 0.3.5b-20260212"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# REPORT
[bool]$HtmlFileReport = $true    # On/Off HTML Report out file
[bool]$CsvFileReport  = $false     # On/Off CSV Report out file
[int]$CsvBufferSize = 20     # Number of lines in buffer to write to CSV file

# LOG
[bool]$WriteLog = $false   # On/Off Logging Function
[bool]$ShowLog = $true   # On/Off console output
[string]$LogDir = "$PSScriptRoot\log"
[string]$LogFile = "$logDir\PsSPM_log_$(Get-Date -Format 'yyyyMMddHHmmss').log"

# SNMP
[int]$SnmpTimeoutMs = 5000  # Timeout SNMP - Milliseconds (1 sec. = 1000 ms) / 0 = infinity
[int]$SnmpMaxAttempts = 3   # Retry Count
[int]$SnmpDelayMs = 2000    # Milliseconds (1 sec. = 1000 ms)

# TCP
[int]$TCPPort = 80         # TCP port for check link
[int]$TcpTimeoutMs = 1000   # Timeout TcpClient - Milliseconds (1 sec. = 1000 ms)
[int]$MaxRetries   = 3     # Retry Count
[int]$RetryDelayMs = 2000  # Milliseconds (1 sec. = 1000 ms)
[int]$TCPThreads = 10      # Ьaximum number of running TCP threads

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

# Printer
[hashtable]$modelPatterns = @{
    "*333*" = "333"
    "*B60*" = "B60"
    "*B61*" = "B60"
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
    "*670*" = "670"
    "*450*" = "450"
    "*C60*" = "C60"
    "*533*" = "B60" # Same as B60
    "*622*" = "622"
    "*611*" = "611"
    "*M42*" = "M42"
    "*M25*" = "M25"
    "*T73*" = "T73"
	"*350*" = "350"
	"*B40*" = "B60"  # Same as B60
}

# Path
[string]$PrinterOIDPath = "$PSScriptRoot\lib\PsSPM_oid.psd1"
[string]$ScriptGuiXaml = "$PSScriptRoot\lib\PsSPM_wpf.ps1"
[string]$HtmlHeader = "$PSScriptRoot\lib\PsSPM_html.ps1"
[string]$MailPath = "$PSScriptRoot\lib\PsSPM_mail.ps1"
[string]$TcpPath = "$PSScriptRoot\lib\PsSPM_tcp.ps1"
[string]$SnmpPath = "$PSScriptRoot\lib\PsSPM_snmp.ps1"
[string]$PrinterListPath = "$PSScriptRoot\ip"
[string]$ReportDir ="$PSScriptRoot\report"
[string]$DllPath = "$PSScriptRoot\lib\SharpSnmpLib.dll"
[string]$selectedFile = $null
[string]$filenamecsv = $null
[string]$filenamehtml = $null

#region Helper Functions
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir | Out-Null }

# Logging function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $logMessage = "$timestamp [$Level] - $Message"
    
    if ($WriteLog) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8
    }
    if ($ShowLog) {
        if ($Level -eq "ERROR") { Write-Host $logMessage -ForegroundColor Red }
        elseif ($Level -eq "WARNING") { Write-Host $logMessage -ForegroundColor Yellow }
        else { Write-Host $logMessage -ForegroundColor Green }
    }
}

function Get-PrinterModelOIDSet {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Model,
        [Parameter(Mandatory=$true)]
        [hashtable]$OIDMapping
    )

    foreach ($pattern in $modelPatterns.Keys) { if ($Model -like $pattern) {return $OIDMapping[$modelPatterns[$pattern]]} }
    return $OIDMapping["Default"]
}

function Get-PrinterData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TargetHost,
        [Parameter(Mandatory=$true)]
        [hashtable]$OidMapping,
        [int]$SnmpTimeoutMs = 5000,
        [int]$SnmpMaxAttempts = 3,
        [int]$SnmpDelayMs = 2000,
        [string]$CurrentPrinter = ""
    )

    $results = [System.Collections.Generic.Dictionary[string, object]]::new()

    try {
        # Get printer model to determine OIDs to use
        $model = Get-SnmpData -Target $TargetHost -Oid $OidMapping["Default"].Model -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs
        if (-not $model) { Write-Log "Failed to get printer model $TargetHost" -Level "WARNING"
            return $null
        }

        Write-Log "$CurrentPrinter : $TargetHost - $($model.result)"
        
        $printername = Get-SnmpData -Target $TargetHost -Oid $OidMapping["Default"].PName -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs
        if (-not $printername) { Write-Log "Failed to get printer name $TargetHost" -Level "WARNING"
            return $null
        }

        $macadr = Get-SnmpData -Target $TargetHost -Oid $OidMapping["Default"].PMac -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs -MacAdr
        if (-not $macadr) { Write-Log "Failed to get printer MAC $TargetHost" -Level "WARNING"
            return $null
        }

        # Determine OID set based on model
        $Script:oidSet = Get-PrinterModelOIDSet -Model $model.result -OIDMapping $OidMapping
        if (-not $oidSet) { Write-Log "Not OID set for model: $($model.result)" -Level "WARNING"
            return $null
        }

        # Safe add to dictionary
        $results['Model']       = $model.result
        $results['PName']       = $printername.result
        $results['IPAddress']   = $TargetHost
        $results['PMac']        = $macadr.result

        foreach ($item in $oidSet.GetEnumerator()) {
            $Value = $null

            if ($item.Name -in @("Display", "Status")) { continue }
            
            if ($item.Name -in @("PMac")) {
                $Name = $item.Name
                try { $Value = Get-SnmpData -Target $TargetHost -Oid $item.Value -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs -MacAdr
                    if ($value.Success -like 'true' ) { $results[$Name] = $Value.result } else { throw }
                }
                catch { Write-Log "Error Get-SnmpData for $Name : $($_.Exception.Message)" -Level "WARNING"
                    $results[$Name] = "Error"
                }
            } else {
                $Name = $item.Name
                try { $Value = Get-SnmpData -Target $TargetHost -Oid $item.Value -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs
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
        'DC'  = @('DrumCUTotal', 'DrumCUCurrent')
        'DM'  = @('DrumMUTotal', 'DrumMUCurrent')
        'DY'  = @('DrumYUTotal', 'DrumYUCurrent')
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
		throw "Missing Lextm.SharpSnmpLib Assembly; is it installed?" | Out-Null
    }

    # Load printer OID
    if (Test-Path $PrinterOIDPath) { $oidMapping = Import-PowerShellDataFile $PrinterOIDPath
        Write-Log "=== OID mapping: OK ==="
    } else { Write-Log "Printer OID file not found: $PrinterOIDPath" -Level "ERROR"
        throw "Printer OID file not found: $PrinterOIDPath" | Out-Null
    }

    # Load Mail
    if ($MailSend) {
        if (Test-Path $MailPath) { . $MailPath
            Write-Log "=== Mail Function: OK ==="
        } else { Write-Log "Mail function file not found: $MailPath" -Level "ERROR"
            throw "Mail function file not found: $MailPath" | Out-Null
        }
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
                Write-Log "=== WPF GUI: OK ==="
            } else { throw " WPF GUI not load" | Out-Null }

            if ($selectedFile = Show-UserGUIXaml -Directory $PrinterListPath) { Write-Log "=== Printer list: OK ===" }
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

    # Test connection
    if (Test-Path $TcpPath) { . $TcpPath
        Write-Log "=== TCP Function: OK ==="
    } else { Write-Log "TCP function file not found: $TcpPath" -Level "ERROR"
        throw "TCP function file not found: $TcpPath" | Out-Null
    }
    Write-Log "=== TCP test connection ==="
    $CheckPrinters = Test-TcpConnectionParallel -Devices $printers.Value -Port $TCPPort -TcpTimeoutMs $TcpTimeoutMs -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs -Threads $TCPThreads

    Write-Log "=== Monitoring $totalPrinters printers ==="

    $DataHtmlReport = [System.Collections.Generic.List[PSObject]]::new()
    $DataCsvReport = [System.Collections.Generic.List[PSObject]]::new($CsvBufferSize)
    if($CsvFileReport) {
        if ($PrinterRange.Value.Count -gt 0) {
            $filenamecsv = "$ReportDir\PsSPM_report_$($PrinterRange[0].Value)-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        }
        else {
            $filePrefixcsv = Split-Path -Path $selectedFile -Leaf
            $filePrefixcsv = $filePrefixcsv -replace '.txt', '' -replace '.csv', ''
            $filenamecsv = "$ReportDir\PsSPM_report_$filePrefixcsv-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        }
    }

    # SNMP
    if (Test-Path $SnmpPath) { . $SnmpPath
        Write-Log "=== SNMP Function: OK ==="
    } else { Write-Log "SNMP function file not found: $SnmpPath" -Level "ERROR"
        throw "SNMP function file not found: $SnmpPath" | Out-Null
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
            Write-Log "$currentPrinter : $printerIP - Offline" -Level "ERROR"
        }
        else {
            try {
                $TcpStatus = "Online"
                $PrinterData = Get-PrinterData -Target $printerIP -OidMapping $OidMapping -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs -CurrentPrinter $CurrentPrinter
                $tonerLevels = Update-TonerLevels -PrinterData $PrinterData

                # Get printer status values
                if ($oidSet.Display) { $pdisplay = Get-SnmpBulkWalkWithEncoding -Target $printerIP -Oid $oidSet.Display -SnmpTimeoutMs $SnmpTimeoutMs }
                if ($oidSet.Status) { $pstatus = Get-SnmpBulkWalkWithEncoding -Target $printerIP -Oid $oidSet.Status -SnmpTimeoutMs $SnmpTimeoutMs }
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
                "<span>Ping</span>" = Format-Status -TcpStatus $TcpStatus
                "<span>Name</span>" = "<a class='printer-link' href='http://$printerIP' target='_blank'>$($PrinterData.PName)</a>"
                "<span>MAC</span>" = "<span style='cursor: copy;' onclick='copyText(this)'>$(Format-Value -Value $PrinterData.PMac)</span>"
                "<span>Model</span>" = Format-Value -Value $PrinterData.Model
                "<span>S/N</span>" = "<span style='cursor: copy;' onclick='copyText(this)'>$(Format-Value -Value $PrinterData.Serial)</span>"
                "<span>Black</span>" = Format-Value -Value $PrinterData.BlackCount
                "<span>Color</span>" = Format-Value -Value $PrinterData.ColorCount
                "<span>Total</span>" = Format-Value -Value $PrinterData.TotalCount
                "<span style='color:#00FFFF'>C</span> Toner"  = if (-not($tonerLevels.TC -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TC)</span>"} else { "" }
                "<span style='color:#FD3DB5'>M</span> Toner"  = if (-not($tonerLevels.TM -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TM)</span>"} else { "" }
                "<span style='color:#FFDE21'>Y</span> Toner"  = if (-not($tonerLevels.TY -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TY)</span>"} else { "" }
                "<span style='color:#000000'>K</span> Toner"  = if (-not($tonerLevels.TK -like $null)) {"<span class='container'>$(Format-TonerLevel -Level $tonerLevels.TK)</span>"} else { "" }
                "<span style='color:#00FFFF'>C</span><span style='color:#FD3DB5'>M</span><span style='color:#FFDE21'>Y</span><span style='color:#000000'>K</span> DrumKit" = `
                    if (-not($tonerLevels.DC -like $null -and $tonerLevels.DM -like $null -and $tonerLevels.DY -like $null -and $tonerLevels.DK -like $null -and $tonerLevels.DKU -like $null)) {
                    "<span class='container'>$(Format-TonerLevel -Level $tonerLevels.DC) $(Format-TonerLevel -Level $tonerLevels.DM) $(Format-TonerLevel -Level $tonerLevels.DY) `
                    $(Format-TonerLevel -Level $tonerLevels.DKU) $(Format-TonerLevel -Level $tonerLevels.DK)</span>"} else { "" }
                "<span style='color:#FFDE21'>Display</span>" = if ($pdisplay.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputDisplay</ul></div>" } else { "" }
                "<span style='color:#FFDE21'>Active Alerts</span>" = if ($pstatus.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputStatus</ul></div>" } else { "" }
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
                "Display" = if ($pdisplay.Length -ne 0) { "$pdisplay" } else { "" }
                "Active Alerts" = if ($pstatus.Length -ne 0) { "$pstatus" } else { "" }
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

        $htmlContent = $DataHtmlReport | ConvertTo-Html -Title "Printer Status Report" -Head $Header -PostContent "<span style='font-size: 12px;'>$Version</span>"
        # Fix HTML encoding
        $htmlContent = $htmlContent -replace '&lt;', '<' -replace '&#39;', "'" -replace '&gt;', '>' -replace'<table>', '<table id="PrinterTable">'

        if ($PrinterRange.Value.Count -gt 0) {
            $filenamehtml = "$ReportDir\PsSPM_report_$($PrinterRange[0].Value)-$(Get-Date -Format 'yyyyMMddHHmmss').html"
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
