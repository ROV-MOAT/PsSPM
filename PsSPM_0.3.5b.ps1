<#
.SYNOPSIS
    PsSPM(ROV-MOAT) - Powershell SNMP Printer Monitoring and Reporting Script
    C# SNMP Library is used - https://github.com/lextudio/sharpsnmplib

.DESCRIPTION
    Checks printer status via SNMP and generates CSV, HTML report with toner levels, counters, and information

.NOTES
    Version: 0.3.5
    Author: Oleg Ryabinin + AI
    Date: 2025-08-22
    
    MESSAGE:
    Powershell 5+

    CHANGELOG:

    Ver. 0.3.5b
    Сode optimization
    
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
[bool]$ShowGUI = $true
[string]$Version = "0.3.5b"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

#REPORT
[bool]$HtmlFileReport = $true    # On/Off HTML Report out file
[bool]$CsvFileReport  = $false     # On/Off CSV Report out file
[int]$CsvBufferSize = 20     # Number of lines in buffer to write to CSV file

#LOG
[bool]$WriteLog = $true   # On/Off Logging Function
$LogDir = "$PSScriptRoot\Log"
$LogFile = "$logDir\PrinterMonitor_$(Get-Date -Format 'yyyyMMddHHmmss').log"

#SNMP
[int]$TimeoutMsUDP = 5000  # Timeout SNMP - Milliseconds (1 sec. = 1000 ms) / 0 = infinity

#TCP
[int]$TCPPort = 80         # TCP port for check link
[int]$TimeoutMsTCP = 500   # Timeout TcpClient - Milliseconds (1 sec. = 1000 ms)
[int]$MaxRetries   = 3     # Retry Count
[int]$RetryDelayMs = 2000  # Milliseconds (1 sec. = 1000 ms)
[int]$TCPThreads = 10      # Ьaximum number of running TCP threads

#Path
$PrinterOIDPath = "$PSScriptRoot\Lib\PsSPM_oid.psd1"
$ScriptGuiXaml = "$PSScriptRoot\Lib\PsSPM_gui_wpf.ps1"
$HtmlHeader = "$PSScriptRoot\Lib\PsSPM_html_header.ps1"
$PrinterListPath = "$PSScriptRoot\IP"
$ReportDir ="$PSScriptRoot\Report"
$DllPath = "$PSScriptRoot\Lib\SharpSnmpLib.dll"
$selectedFile = $null

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
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
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
                    $client = New-Object System.Net.Sockets.TcpClient
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

        # Set up SNMP manager
        $IP = [System.Net.IPAddress]::Parse($Target)
        $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)
        $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
        $null = $vList.Add([Lextm.SharpSnmpLib.Variable]::new([Lextm.SharpSnmpLib.ObjectIdentifier]::new($Oid)))

    try {
        $result = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            [Lextm.SharpSnmpLib.OctetString]::new($Community),
            $vList,
            $TimeoutMsUDP
        )
        return $result.Data.ToString()
    }
    catch { Write-Log "SNMP query failed for $Target (OID: $Oid): $_" -Level "WARNING" 
        return "Failed"
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
        # Set up SNMP manager
        $IP = [System.Net.IPAddress]::Parse($Target)
        $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)
        $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
        $result_out = [System.Collections.Generic.List[PSObject]]::new()
        $encodings = @([System.Text.Encoding]::UTF8, [System.Text.Encoding]::ASCII, [System.Text.Encoding]::GetEncoding(1251))

    try {
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
    catch { Write-Log "SNMP Walk query failed for $Target (OID: $Oid): $_" -Level "WARNING"; return $null }
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

    foreach ($pattern in $modelPatterns.Keys) { if ($Model -like $pattern) { return $OIDMapping[$modelPatterns[$pattern]] } }
    return $OIDMapping["Default"]
}

function Get-PrinterData {
    param ([string]$TargetHost)

    $results = [System.Collections.Generic.Dictionary[string, object]]::new()

    # Get printer model to determine OIDs to use
    $model = Get-SnmpData -Target $printerIP -Oid $oidMapping["Default"].Model -TimeoutMsUDP $TimeoutMsUDP

    Write-Host "$currentPrinter : $printerIP - $model" -ForegroundColor Green
    Write-Log "Checking printer: $printerIP - Online"
    $printername = Get-SnmpData -Target $printerIP -Oid $oidMapping["Default"].PName -TimeoutMsUDP $TimeoutMsUDP

    # Determine OID set based on model
    $Script:oidSet = Get-PrinterModelOIDSet -Model $model -OIDMapping $oidMapping

    $results.Add('Model', $model)
    $results.Add('PName', $printername)

    foreach ($item in $oidSet.GetEnumerator()) {
        if ($item.Name -like "Display" -or $item.Name -like "Status") { continue }
        elseif ($item.Value -is [string]) {
            $Name = $item.Name
            $Value = Get-SnmpData -Target $TargetHost -Oid $item.Value -TimeoutMsUDP $TimeoutMsUDP
            $results.Add($Name, $Value)
        }
    }
    return $results
}

function Get-TonerPercentage {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Total,
        [Parameter(Mandatory=$true)]
        [object]$Current
    )
    
    if ($Total -and $Total -notlike 'Failed' -and [int]$Total -gt 0) {
        if ($Current -and $Current -notlike 'Failed' -and [int]$Current -gt 0) { return [math]::Round(([int]$Current / [int]$Total) * 100) }
        elseif ($Total -like 'Failed' -or $Current -like 'Failed') { return "Failed" }
        else { return 0 }
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

function Format-Value {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Value,
        [string]$Default = "Error"
    )
    if ($Value -and $Value -notlike 'Failed') { return "<center>$Value</center>" } elseif ($Value -like 'Failed') { return "<center class='error'>$Default</center>" } else { return "" }
}

function Format-TonerLevel {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Level
    )

    if ($Level -notlike 'Failed' -and [int]$Level -gt 49) { return "<div class='container'><center><span class='toner-high'>$Level%</span></center></div>" }
    elseif ($Level -notlike 'Failed' -and [int]$Level -gt 10 -and [int]$Level -le 49) { return "<div class='container'><center><span class='toner-medium'>$Level%</span></center></div>" }
    elseif ($Level -notlike 'Failed' -and $Level -notlike "" -and [int]$Level -le 10) { return "<div class='container'><center><span class='toner-low'>$Level%</span></center></div>" }
    elseif ($Level -like 'Failed') { return "<div class='container'><center><span class='error'>Error</span></center></div>" }
    else { return "" }
}
#endregion

#region Initialization
try {
    Write-Log "=== Starting Printer Monitoring Version: $Version ==="

    # Load SharpSNMPLib.dll
    if ([System.Reflection.Assembly]::LoadFrom($DllPath)) { Write-Log "=== SharpSnmpLib: OK ===" }
    else {
        Write-Log "Failed to load Lextm.SharpSnmpLib" -Level "ERROR"
		throw "Missing Lextm.SharpSnmpLib Assembly; is it installed?" | Out-Null
    }

    # Load printer OID
    if (Test-Path $PrinterOIDPath) {
        $oidMapping = Import-PowerShellDataFile $PrinterOIDPath
        Write-Log "=== OID mapping: OK ==="
    } else {
        Write-Log "Printer OID file not found: $PrinterOIDPath" -Level "ERROR"
        throw "Printer OID file not found: $PrinterOIDPath" | Out-Null
    }

    # Load GUI
    if (Test-Path $ScriptGuiXaml) {
        # File exists, dot source it
        . $ScriptGuiXaml
        Write-Log "=== GUI Xaml: OK ==="
    } else { throw "GUI Xaml not load" | Out-Null }

    # Load printer list
    $PrinterRange = @()
    if ($ShowGUI) {
        if ($selectedFile = Show-UserGUIXaml -Directory $PrinterListPath) { Write-Log "=== Printer list: OK ===" }
        else {
            Write-Log "Printer list file not load" -Level "ERROR"
            throw "Printer list file not load" | Out-Null
        }
    } else {
        if ($selectedFile = Open-File $PrinterListPath) { Write-Log "=== Printer list: OK ===" }
        else {
            Write-Log "Printer list file not load" -Level "ERROR"
            throw "Printer list file not load" | Out-Null
        }
    }
    
    # Import from (range/txt/csv)
    if ($PrinterRange -notlike $null) { $printers = $PrinterRange }
    elseif ($printers = Import-Csv -Path $selectedFile -Header Value) { Write-Log "=== Import Printers IP: OK ===" }
    else {
        Write-Log "=== Import Printers IP: Error ===" -Level "ERROR"
        throw "Import Printers IP: Error" | Out-Null
    }

    $totalPrinters = $printers.Value.Count
    
    if ($totalPrinters -eq 0) {
        Write-Log "No printers found in the list" -Level "ERROR"
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
        $OnlineDevice = $currentPrinterIP.Connected
        $TcpStatus = $null
        $oidSet = $null
        $pdisplay = $null
        $pstatus = $null
        $PrinterData = $null
        $tonerLevels = @{}

        Write-Log "=== Checking printer: $currentPrinter of $totalPrinters ==="
        # Check printer availability(Ping)
        if ($OnlineDevice -like 'false') {
            $TcpStatus = "Offline"
            Write-Host "$currentPrinter : $printerIP - Offline" -ForegroundColor Red
            Write-Log "Checking printer: $printerIP - Offline"
        }
        else {
            try {
                $TcpStatus = "Online"
                $PrinterData = Get-PrinterData -Target $printerIP

                if ($PrinterData.ContainsKey('TonerKTotal')) { $tonerLevels.TK = Get-TonerPercentage -Total $PrinterData.TonerKTotal -Current $PrinterData.TonerKCurrent }
                if ($PrinterData.ContainsKey('DrumKUTotal')) { $tonerLevels.DKU = Get-TonerPercentage -Total $PrinterData.DrumKUTotal -Current $PrinterData.DrumKUCurrent }
                if ($PrinterData.ContainsKey('DrumKTotal')) { $tonerLevels.DK = Get-TonerPercentage -Total $PrinterData.DrumKTotal -Current $PrinterData.DrumKCurrent }
                
                if ($PrinterData.ContainsKey('TonerCTotal')) {
                    $tonerLevels.TC = Get-TonerPercentage -Total $PrinterData.TonerCTotal -Current $PrinterData.TonerCCurrent
                    $tonerLevels.TM = Get-TonerPercentage -Total $PrinterData.TonerMTotal -Current $PrinterData.TonerMCurrent
                    $tonerLevels.TY = Get-TonerPercentage -Total $PrinterData.TonerYTotal -Current $PrinterData.TonerYCurrent
                }

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
            $htmlD = foreach ($element in $pdisplay) { "<li>$element</li>" | Out-String -Stream }
            $htmlOutputD = $htmlD -join ""

            $htmlS = foreach ($element in $pstatus) { "<li>$element</li>" | Out-String -Stream }
            $htmlOutputS = $htmlS -join ""

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
                "<span style='color:#FFDE21'>Display</span>" = if ($pdisplay.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputD</ul></div>" } else { "" }
                "<span style='color:#FFDE21'>Active Alerts</span>" = if ($pstatus.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputS</ul></div>" } else { "" }
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
    #HTML
    If($HtmlFileReport) {
        if (Test-Path $HtmlHeader) {
            . $HtmlHeader
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
            Start-Process $filenamehtml
        }
        else { Write-Log "Failed to generate HTML report" -Level "ERROR" }
    }
    
    #CSV
    if($CsvFileReport) {
        if (Test-Path $filenamecsv) {
            Write-Log "Report generated: $filenamecsv"
            Start-Process $filenamecsv
        }
        else { Write-Log "Failed to generate CSV report" -Level "ERROR" }
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