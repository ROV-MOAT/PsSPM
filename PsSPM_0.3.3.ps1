<#
.SYNOPSIS
    PsSPM(ROV-MOAT) - Powershell SNMP Printer Monitoring and Reporting Script
    C# SNMP Library is used - https://github.com/lextudio/sharpsnmplib

.DESCRIPTION
    Checks printer status via SNMP and generates CSV, HTML report with toner levels, counters, and information

.NOTES
    Version: 0.3.3
    Author: Oleg Ryabinin + AI
    Date: 2025-08-07
    
    MESSAGE:
    Powershell 5+

    CHANGELOG:

    Ver. 0.3.3
    + RunspacePool (Test-TcpConnectionParallel)

    Ver. 0.3.2
    Visual changes in HTML report
    + GUI
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
$ShowGUI = $true

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PrinterOIDPath = "$PSScriptRoot\Lib\PsSPM_oid.psd1"
$PrinterListPath = "$PSScriptRoot\IP"
$ReportDir ="$PSScriptRoot\Report"
$DllPath = "$PSScriptRoot\Lib\SharpSnmpLib.dll"
$selectedFile = $null

#REPORT
$HtmlFileReport = $true    # On/Off HTML Report out file
$CsvFileReport  = $false     # On/Off CSV Report out file
$CsvBufferSize = 20     # Number of lines in buffer to write to CSV file

#LOG
$WriteLog = $false   # On/Off Logging Function
$LogDir = "$PSScriptRoot\Log"
$LogFile = "$logDir\PrinterMonitor_$(Get-Date -Format 'yyyyMMddHHmmss').log"

#SNMP
[int]$TimeoutMsUDP = 5000  # Timeout SNMP - Milliseconds (1 sec. = 1000 ms) / 0 = infinity

#TCP
[int]$TimeoutMsTCP = 500   # Timeout TcpClient - Milliseconds (1 sec. = 1000 ms)
[int]$MaxRetries   = 3     # Retry Count
[int]$RetryDelayMs = 2000  # Milliseconds (1 sec. = 1000 ms)
[int]$TCPThreads = 10      # Ьaximum number of running TCP threads

#region HTML Template
$Header = @"
<style>
    body { font-family: 'Trebuchet MS', sans-serif; margin: 20px; }
    table {
        border-collapse: collapse;
        border: 1px solid black;
        font-size: 90%;
        width: 100%;
        margin-bottom: 20px;
    }
    th, td {
        border: 2px solid #ddd;
        padding: 5px;
        text-align: center;
        font-weight: normal;
        font-size: 15px;
    }
    th {
        font-size: 17px;
        background-color: #6d8196;
        color: white;
        position: sticky;
        top: 0;
    }
    tr:hover { background-color: #f0f0f0; }
    li { margin-top: 5px; margin-bottom: 5px; }
    .online { color: green; }
    .offline { color: red; }
    .error { color: orange; }

    .toner-high { color: green; font-weight: bold; transition: all 0.3s ease; }
    .toner-medium { color: orange; font-weight: bold; transition: all 0.3s ease; }
    .toner-low { color: red; font-weight: bold; transition: all 0.3s ease; }

    a.printer-link { text-decoration: none; color: #0066cc; }
    a.show-link {
        display: inline-block;
        text-decoration: none;
        color: #0066cc;
        animation: show-link ease-in-out 1s infinite alternate;
    }
    @keyframes show-link {
        0% {
            transform: rotate(0deg);
        }
        25% {
            transform: rotate(5deg);
        }
        75% {
            transform: rotate(-5deg);
        }
        100% {
            transform: rotate(0deg);
        }
    }

    .tooltip { position: relative; display: inline-block; }

    /* Tooltip text */
    .tooltip .tooltiptext {
        list-style-position: inside;
        list-style-type: disclosure-closed;
        visibility: hidden;
        max-width: 400px;
        width: max-content; /* Allows the tooltip to size based on content up to max-width */
        white-space: normal; /* Ensures text wraps within the tooltip */
        word-wrap: break-word; /* Prevents long words from overflowing */
        background-color: #6d8196;
        color: #ffffff;
        text-align: Left;
        padding: 5px 5px 7px 5px;
        margin: 0;
        border-radius: 5px;
        opacity: 0;
        transition: visibility 0.3s ease, opacity 0.3s ease, background-color 0.3s ease;
 
        /* Position the tooltip text */
        position: absolute;
        z-index: 1000;
        right: 35px;
        top: 20px;
    }

    /* Show the tooltip text when you mouse over the tooltip container */
    .tooltip:hover .tooltiptext { visibility: visible; opacity: 1; }
    .tooltiptext:hover { background-color: #000000; }

    .container { border-radius: 3px; padding: 5px; margin: 0; }
    .container:hover .toner-high,
    .container:hover .toner-medium,
    .container:hover .toner-low {
        background-color: #000000; /* Цвет фона */
        border-radius: 4px; /* Закругление углов */
        padding: 2px 3px 2px 3px; /* Отступы вокруг текста - верх, право, низ, лево */
        color: #ffffff;
    }
</style>
"@
#endregion

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

#GUI
function Show-UserGUI ([string] $initialDirectory) {

    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PsSPM 0.3.3"
    $form.Size = New-Object System.Drawing.Size(290, 335)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.TopMost = $true
    $groupBox1 = New-Object System.Windows.Forms.GroupBox
    $groupBox2 = New-Object System.Windows.Forms.GroupBox
    $checkBox1 = New-Object System.Windows.Forms.CheckBox
    $checkBox2 = New-Object System.Windows.Forms.CheckBox
    $Labelcsvbuf = New-Object System.Windows.Forms.Label
    $Labelsnmp = New-Object System.Windows.Forms.Label
    $LabelTcp = New-Object System.Windows.Forms.Label
    $LabelTcpMaxRetries = New-Object System.Windows.Forms.Label
    $LabelTcpRetryDelay = New-Object System.Windows.Forms.Label
    $LabelTcpThreads = New-Object System.Windows.Forms.Label
    $numericUpDownCsvBuf = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownSnmp = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcp = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcpMaxRetries = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcpRetryDelay = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcpThreads = New-Object System.Windows.Forms.NumericUpDown
    $form.Controls.Add($groupbox1)
    $form.Controls.Add($groupBox2)
    
    # GroupBox 1
    $groupBox1.Location = New-Object System.Drawing.Point(15,10)
    $groupBox1.size = New-Object System.Drawing.Size(110,130)
    $groupBox1.text = "Report"
    $groupBox1.Visible = $true
    $groupbox1.Controls.Add($checkBox1)
    $groupbox1.Controls.Add($checkBox2)
    $groupbox1.Controls.Add($Labelcsvbuf)
    $groupbox1.Controls.Add($numericUpDownCsvBuf)

    # GroupBox 2
    $groupBox2.Location = New-Object System.Drawing.Point(150,10)
    $groupBox2.size = New-Object System.Drawing.Size(110,270)
    $groupBox2.text = "TCP/SNMP"
    $groupBox2.Visible = $true
    $groupbox2.Controls.Add($Labelsnmp)
    $groupbox2.Controls.Add($numericUpDownSnmp)
    $groupbox2.Controls.Add($LabelTcp)
    $groupbox2.Controls.Add($numericUpDownTcp)
    $groupbox2.Controls.Add($LabelTcpMaxRetries)
    $groupbox2.Controls.Add($numericUpDownTcpMaxRetries)
    $groupbox2.Controls.Add($LabelTcpRetryDelay)
    $groupbox2.Controls.Add($numericUpDownTcpRetryDelay)
    $groupbox2.Controls.Add($LabelTcpThreads)
    $groupbox2.Controls.Add($numericUpDownTcpThreads)

    # CheckBox 1
    $checkBox1.Location = New-Object System.Drawing.Point(10,20)
    $checkBox1.Size = New-Object System.Drawing.Size(95,20)
    $checkBox1.Text = "HTML Report"
    $checkBox1.Checked = $script:HtmlFileReport
    $checkBox1.Add_CheckedChanged({ $script:HtmlFileReport = $checkBox1.Checked })

    # CheckBox 2
    $checkBox2.Location = New-Object System.Drawing.Point(10,44)
    $checkBox2.Size = New-Object System.Drawing.Size(95,20)
    $checkBox2.Text = "CSV Report"
    $checkBox2.Checked = $script:CsvFileReport
    $checkBox2.Add_CheckedChanged({ $script:CsvFileReport = $checkBox2.Checked })

    # CheckBox 3
    $checkBox3 = New-Object System.Windows.Forms.CheckBox
    $checkBox3.Location = New-Object System.Drawing.Point(20,145)
    $checkBox3.Size = New-Object System.Drawing.Size(95,20)
    $checkBox3.Text = "Log enable"
    $checkBox3.Checked = $script:WriteLog
    $checkBox3.Add_CheckedChanged({ $script:WriteLog = $checkBox3.Checked })
    $form.Controls.Add($checkBox3)

    # Label CsvBufferSize
    $Labelcsvbuf.Text = "CSV Buffer Size:"
    $Labelcsvbuf.Location = New-Object System.Drawing.Point(10, 80)
    $Labelcsvbuf.AutoSize = $true

    # Create a CsvBufferSize control
    $numericUpDownCsvBuf.Location = New-Object System.Drawing.Point(10, 100)
    $numericUpDownCsvBuf.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownCsvBuf.Minimum = 5
    $numericUpDownCsvBuf.Maximum = 500
    $numericUpDownCsvBuf.Value = $CsvBufferSize
    $numericUpDownCsvBuf.Increment = 5

    # Label SNMP timeout
    $Labelsnmp.Text = "SNMP timeout ms:"
    $Labelsnmp.Location = New-Object System.Drawing.Point(10, 20)
    $Labelsnmp.AutoSize = $true

    # Create a SNMP control
    $numericUpDownSnmp.Location = New-Object System.Drawing.Point(10, 40)
    $numericUpDownSnmp.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownSnmp.Minimum = 50
    $numericUpDownSnmp.Maximum = 10000
    $numericUpDownSnmp.Value = $TimeoutMsUDP
    $numericUpDownSnmp.Increment = 50

    # Label TCP timeout
    $LabelTcp.Text = "TCP timeout ms:"
    $LabelTcp.Location = New-Object System.Drawing.Point(10, 70)
    $LabelTcp.AutoSize = $true

    # Create a TCP control
    $numericUpDownTcp.Location = New-Object System.Drawing.Point(10, 90)
    $numericUpDownTcp.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcp.Minimum = 50
    $numericUpDownTcp.Maximum = 10000
    $numericUpDownTcp.Value = $TimeoutMsTCP
    $numericUpDownTcp.Increment = 50

    # Label TCP retry
    $LabelTcpMaxRetries.Text = "Max Retries:"
    $LabelTcpMaxRetries.Location = New-Object System.Drawing.Point(10, 120)
    $LabelTcpMaxRetries.AutoSize = $true

    # Create a TCP retry control
    $numericUpDownTcpMaxRetries.Location = New-Object System.Drawing.Point(10, 140)
    $numericUpDownTcpMaxRetries.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcpMaxRetries.Minimum = 1
    $numericUpDownTcpMaxRetries.Maximum = 20
    $numericUpDownTcpMaxRetries.Value = $MaxRetries
    $numericUpDownTcpMaxRetries.Increment = 1

    # Label TCP Retry Delay
    $LabelTcpRetryDelay.Text = "Retry Delay ms:"
    $LabelTcpRetryDelay.Location = New-Object System.Drawing.Point(10, 170)
    $LabelTcpRetryDelay.AutoSize = $true

    # Create a TCP Retry Delay control
    $numericUpDownTcpRetryDelay.Location = New-Object System.Drawing.Point(10, 190)
    $numericUpDownTcpRetryDelay.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcpRetryDelay.Minimum = 50
    $numericUpDownTcpRetryDelay.Maximum = 10000
    $numericUpDownTcpRetryDelay.Value = $RetryDelayMs
    $numericUpDownTcpRetryDelay.Increment = 50

    # Label TCP Threads
    $LabelTcpThreads.Text = "TCP Threads:"
    $LabelTcpThreads.Location = New-Object System.Drawing.Point(10, 220)
    $LabelTcpThreads.AutoSize = $true

    # Create a TCP Threads control
    $numericUpDownTcpThreads.Location = New-Object System.Drawing.Point(10, 240)
    $numericUpDownTcpThreads.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcpThreads.Minimum = 5
    $numericUpDownTcpThreads.Maximum = 1000
    $numericUpDownTcpThreads.Value = $TCPThreads
    $numericUpDownTcpThreads.Increment = 5

    # Label
    $Labelsf = New-Object System.Windows.Forms.Label
    $Labelsf.Text = "Select file:"
    $Labelsf.Location = New-Object System.Drawing.Point(15, 170)
    $Labelsf.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $Labelsf.AutoSize = $true
    $Form.Controls.Add($Labelsf)

    # ComboBox
    $files = Get-ChildItem -Path $InitialDirectory -File | Where-Object { $_.Extension -eq ".txt" -or $_.Extension -eq ".csv" } | Select-Object -ExpandProperty Name
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(15, 190)
    $comboBox.Size = New-Object System.Drawing.Size(130, 30)
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboBox.Anchor = 'Top, Left, Right'
    $comboBox.Items.AddRange($files)
    $form.Controls.Add($comboBox)

    # Button
    $Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Point(25,240)
    $Button.Size = New-Object System.Drawing.Size(100,30)
    $Button.Text = "Run"
    $Button.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $Button.Add_Click({
        if ($comboBox.SelectedItem -notlike $null) {
                $script:selectedFile = Join-Path -Path $PrinterListPath -ChildPath $comboBox.SelectedItem
                $script:CsvBufferSize = $numericUpDownCsvBuf.Value
                $script:TimeoutMsUDP = $numericUpDownSnmp.Value
                $script:TimeoutMsTCP = $numericUpDownTcp.Value
                $script:MaxRetries = $numericUpDownTcpMaxRetries.Value
                $script:RetryDelayMs = $numericUpDownTcpRetryDelay.Value
                $script:TcpThreads = $numericUpDownTcpThreads.Value
                $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a file.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $form.Controls.Add($Button)

    $form.ShowDialog() | Out-Null

    return $script:selectedFile
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
    $totalTasks = $runspaces.Count
    Write-Log "All $totalTasks tests started, waiting for completion..."
    $completedCount = 0

    while ($runspaces | Where-Object { -not $_.Handle.IsCompleted }) {
        $completedCount = ($runspaces | Where-Object { $_.Handle.IsCompleted }).Count
        $total = $runspaces.Count
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
        $endpoint = New-Object System.Net.IPEndPoint ($IP, $UDPport)
        $communityObject = New-Object Lextm.SharpSnmpLib.OctetString ($Community)
        $vList = New-Object 'System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]'
        $null = $vList.Add((New-Object Lextm.SharpSnmpLib.Variable -ArgumentList @(New-Object Lextm.SharpSnmpLib.ObjectIdentifier ($Oid))))
        $result = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityObject,
            $vList,
            $TimeoutMsUDP
        )
        return $result.Data.ToString()
    }
    catch { Write-Log "SNMP query failed for $Target (OID: $Oid): $_" -Level "WARNING" }
}

function Get-SnmpWalkWithEncoding {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [string]$Community = "public",
        [Parameter(Mandatory=$true)]
        [string]$Oid,
        [int]$UDPport = 161,
        [int]$TimeoutMsUDP = 5000
    )

    # Table to store results
    try {
        # Set up SNMP manager
        $IP = [System.Net.IPAddress]::Parse($Target)
        $endpoint = New-Object System.Net.IPEndPoint ($IP, $UDPport)
        $communityObject = New-Object Lextm.SharpSnmpLib.OctetString ($Community)
        $Oids = New-Object Lextm.SharpSnmpLib.ObjectIdentifier ($Oid)
        $vList = New-Object 'System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]'

        # Create walk mode
        $walkMode = [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree

        # Perform the walk
        $null = [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityObject,
            $Oids,
            $vList,
            $TimeoutMsUDP,
            $walkMode
        )

        $result_out = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($result in $vList) {
            $data = $result.Data
            $value = $null

            # Handle OctetString encoding
            if ($data.TypeCode -eq [Lextm.SharpSnmpLib.SnmpType]::OctetString) {
                try {
                    # Try UTF-8 first
                    $value = $data.ToString([System.Text.Encoding]::UTF8)
                    
                    # If it contains replacement characters, try ASCII
                    if ($value -match '�') {
                        $value = $data.ToString([System.Text.Encoding]::ASCII)
                    }
                }
                catch { $value = $data.ToString() }

                $null = $result_out.Add([Environment]::NewLine+"$value")
            }
        }
        $cleanArray = $result_out.Where( {$_.Trim() -ne ""} )
        return $cleanArray
    }
    catch { Write-Log "SNMP Walk query failed for $Target (OID: $Oid): $_" -Level "WARNING" }
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

    foreach ($pattern in $modelPatterns.Keys) {
        if ($Model -like $pattern) { return $OIDMapping[$modelPatterns[$pattern]] }
    }
    return $OIDMapping["Default"]
}

function Get-TonerPercentage {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Target,
        [Parameter(Mandatory=$true)]
        [string]$TotalOID,
        [Parameter(Mandatory=$true)]
        [string]$CurrentOID
    )

    $total = Get-SnmpData -Target $Target -Oid $TotalOID
    $current = Get-SnmpData -Target $Target -Oid $CurrentOID

    $TotalInt = [Convert]::ToInt32($total.ToString())
    $CurrentInt = [Convert]::ToInt32($current.ToString())
    
    if ($TotalInt -gt 0) { return [math]::Round(($CurrentInt / $TotalInt) * 100) } else { return 0 }
}

function Format-Status {
    param([string]$Status)
    
    $class = switch ($Tcpstatus) {
        "Online"  { "online" }
        "Offline" { "offline" }
        default   { "error" }
    }
    return "<center><span class='$class'>$Tcpstatus</span></center>"
}

function Format-Value {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Value,
        [string]$Default = "Error"
    )
    
    if ($Value -and $Value -gt 0) { return "<center>$Value</center>" } else { return "<center class='error'>$Default</center>" }
}

function Format-TonerLevel {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [int]$Level
    )
    
    $class = switch ($Level) {
        { $_ -gt 49 }  { "toner-high" }
        { $_ -gt 10 -and $_ -le 49 }  { "toner-medium" }
        default       { "toner-low" }
    }
    return "<div class='container'><center><span class='$class'>$Level%</span></center></div>"
}
#endregion

#region Initialization
try {
    Write-Log "=== Starting Printer Monitoring Version: 0.3.3 ==="

    # Load SharpSNMPLib.dll
    if (-not (Add-Type -Path $DllPath -PassThru)) {
        Write-Log "Failed to load Lextm.SharpSnmpLib" -Level "ERROR"
		throw "Missing Lextm.SharpSnmpLib Assembly; is it installed?" | Out-Null
	} else { Write-Log "=== SharpSnmpLib: OK ===" }

    # Load printer OID
    if (-not (Test-Path $PrinterOIDPath)) {
        Write-Log "Printer OID file not found: $PrinterOIDPath" -Level "ERROR"
        throw "Printer OID file not found: $PrinterOIDPath" | Out-Null
    } else {
        $oidMapping = Import-PowerShellDataFile $PrinterOIDPath
        Write-Log "=== OID mapping: OK ==="
    }

    # Load printer list
    if (-not $ShowGUI) {
        if (-not ($selectedFile = Open-File $PrinterListPath)) {
            Write-Log "Printer list file not load" -Level "ERROR"
            throw "Printer list file not load" | Out-Null
        } else { Write-Log "=== Printer list: OK ===" }
    } else { 
        if (-not ($selectedFile = Show-UserGUI $PrinterListPath)) {
            Write-Log "Printer list file not load" -Level "ERROR"
            throw "Printer list file not load" | Out-Null
        } else { Write-Log "=== Printer list: OK ===" }
    }
    
    # Import from file (txt/csv)
    if (-not ($printers = Import-Csv -Path $selectedFile -Header Value)) {
        Write-Log "=== Import-CSV: Error ===" -Level "ERROR"
        throw "Import-CSV: Error" | Out-Null
    } else { Write-Log "=== Import-CSV: OK ===" }

    $totalPrinters = $printers.Value.Count
    
    if ($totalPrinters -eq 0) {
        Write-Log "No printers found in the list" -Level "ERROR"
        throw "No printers found in the list" | Out-Null
    }
    
    Write-Log "=== TCP test connection ==="
    $CheckPrinters = Test-TcpConnectionParallel -Devices $printers.Value -Port 80 -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs -TimeoutMsTCP $TimeoutMsTCP -Threads $TCPThreads

    Write-Log "=== Monitoring $totalPrinters printers ==="

    $DataHtmlReport = [System.Collections.Generic.List[PSObject]]::new()
    $DataCsvReport = [System.Collections.Generic.List[PSObject]]::new($CsvBufferSize)
    if($CsvFileReport) {
        $filePrefixcsv = Split-Path -Path $selectedFile -Leaf
        $filePrefixcsv = $filePrefixcsv -replace '.txt', '' -replace '.csv', ''
        $filenamecsv = "$ReportDir\Printers_report_$filePrefixcsv $(Get-Date -Format 'yyyyMMddHHmmss').csv"
        $DataCsvReport | Export-Csv -Path $filenamecsv -NoTypeInformation -Encoding UTF8 -Delimiter ';' 
    }
#endregion

#region Main Processing
    $start = Get-Date
    $currentPrinter = 0

    ForEach ($currentPrinterIP in $CheckPrinters) {
        $currentPrinter++
        $printerIP = $currentPrinterIP.Device
        $OnlineDevice = $currentPrinterIP.Connected
        $Tcpstatus = $null
        $model = $null
        $printername = $null
        $serial = $null
        $pdisplay = $null
        $pstatus = $null
        $oidSet = $null
        $counters = @{}
        $tonerLevels = @{}

        Write-Log "=== Checking printer: $currentPrinter of $totalPrinters ==="
        # Check printer availability(Ping)
        if ($OnlineDevice -like 'false') {
            Write-Host "$currentPrinter : $printerIP - Offline" -ForegroundColor Red
            Write-Log "Checking printer: $printerIP - Offline"
            $Tcpstatus = "Offline"
        }
        else {
            try {
                # Get printer model to determine OIDs to use
                $model = Get-SnmpData -Target $printerIP -Oid $oidMapping["Default"].Model
                $Tcpstatus = "Online"
                Write-Host "$currentPrinter : $printerIP - $model" -ForegroundColor Green
                Write-Log "Checking printer: $printerIP - Online"
                $printername = Get-SnmpData -Target $printerIP -Oid $oidMapping["Default"].PName

                # Determine OID set based on model
                $oidSet = Get-PrinterModelOIDSet -Model $model -OIDMapping $oidMapping

                # Get serial number
                if ($oidSet.Serial) { $serial = Get-SnmpData -Target $printerIP -Oid $oidSet.Serial }

                # Get counter values
                if ($oidSet.BlackCount) { $counters.Black = Get-SnmpData -Target $printerIP -Oid $oidSet.BlackCount }
                if ($oidSet.ColorCount) { $counters.Color = Get-SnmpData -Target $printerIP -Oid $oidSet.ColorCount }
                if ($oidSet.TotalCount) { $counters.Total = Get-SnmpData -Target $printerIP -Oid $oidSet.TotalCount }
                
                # Get toner levels
                if ($oidSet.TonerK) { $tonerLevels.TK = Get-TonerPercentage -Target $printerIP -TotalOID $oidSet.TonerK.Total -CurrentOID $oidSet.TonerK.Current }

                if ($oidSet.DrumKU) { $tonerLevels.DKU = Get-TonerPercentage -Target $printerIP -TotalOID $oidSet.DrumKU.Total -CurrentOID $oidSet.DrumKU.Current }

                # Get CMYK DrumKit levels
                if ($oidSet.DrumK) { $tonerLevels.DK = Get-TonerPercentage -Target $printerIP -TotalOID $oidSet.DrumK.Total -CurrentOID $oidSet.DrumK.Current }

                # Get toner levels for color printers
                if ($oidSet.TonerC) {
                    $tonerLevels.TC = Get-TonerPercentage -Target $printerIP -TotalOID $oidSet.TonerC.Total -CurrentOID $oidSet.TonerC.Current
                    $tonerLevels.TM = Get-TonerPercentage -Target $printerIP -TotalOID $oidSet.TonerM.Total -CurrentOID $oidSet.TonerM.Current
                    $tonerLevels.TY = Get-TonerPercentage -Target $printerIP -TotalOID $oidSet.TonerY.Total -CurrentOID $oidSet.TonerY.Current
                    # Add other colors similarly
                }

                # Get printer status values
                if ($oidSet.Display) { $pdisplay = Get-SnmpWalkWithEncoding -Target $printerIP -Oid $oidSet.Display }
                if ($oidSet.Status) { $pstatus = Get-SnmpWalkWithEncoding -Target $printerIP -Oid $oidSet.Status }
            }
            catch {
                Write-Log "Error querying $printerIP : $_" -Level "ERROR"
                $Tcpstatus = "Error"
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
                "<span>Name</span>" = "<a class='printer-link' href='http://$printerIP' target='_blank'>$printername</a>"
                "<span>Ping</span>" = Format-Status -Status $Tcpstatus
                "<span>Model</span>" = Format-Value -Value $model -Default ""
                "<span>S/N</span>" = Format-Value -Value $serial -Default ""
                "<span>Black</span>" = Format-Value -Value $counters.Black -Default ""
                "<span>Color</span>" = Format-Value -Value $counters.Color -Default ""
                "<span>Total</span>" = Format-Value -Value $counters.Total -Default ""
                "<span style='color:#00FFFF'>C</span> Toner"  = if ($tonerLevels.TC -ge 0) { Format-TonerLevel -Level $tonerLevels.TC } else { "" }
                "<span style='color:#FD3DB5'>M</span> Toner"  = if ($tonerLevels.TM -ge 0) { Format-TonerLevel -Level $tonerLevels.TM } else { "" }
                "<span style='color:#FFDE21'>Y</span> Toner"  = if ($tonerLevels.TY -ge 0) { Format-TonerLevel -Level $tonerLevels.TY } else { "" }
                "<span style='color:#000000'>K</span> Toner"  = if ($tonerLevels.TK -ge 0) { Format-TonerLevel -Level $tonerLevels.TK } else { "" }
                "<span style='color:#000000'>K</span> Drum"   = if ($tonerLevels.DKU -ge 0) { Format-TonerLevel -Level $tonerLevels.DKU } else { "" }
                "<span style='color:#00FFFF'>C</span><span style='color:#FD3DB5'>M</span><span style='color:#FFDE21'>Y</span><span style='color:#000000'>K</span> DrumKit" = if ($tonerLevels.DK -ge 0) { Format-TonerLevel -Level $tonerLevels.DK } else { "" }
                "<span style='color:#FFDE21'>Display</span>" = if ($pdisplay.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputD</ul></div>" } else { "" }
                "<span style='color:#FFDE21'>Active Alerts</span>" = if ($pstatus.Length -ne 0) { "<div class='tooltip'><a class='show-link' href='http://$printerIP' target='_blank'>Show</a><ul class='tooltiptext'>$htmlOutputS</ul></div>" } else { "" }
                # Add other columns similarly
            })
        }

        if($CsvFileReport) {
            # Add information to CSV report
            $null = $DataCsvReport.Add([PSCustomObject]@{
                "IP"    = "$printerIP"
                "Name"  = "$printername"
                "Ping"  = "$Tcpstatus"
                "Model" = "$model"
                "S/N"   = "$serial"
                "Black" = ($counters.Black)
                "Color" = ($counters.Color)
                "Total" = ($counters.Total)
                "C Toner %" = ($tonerLevels.TC)
                "M Toner %" = ($tonerLevels.TM)
                "Y Toner %" = ($tonerLevels.TY)
                "K Toner %" = ($tonerLevels.TK)
                "K Drum %"  = ($tonerLevels.DKU)
                "CMYK DrumKit %" = ($tonerLevels.DK)
                #"Display" = if ($pdisplay.Length -ne 0) { "$pdisplay" } else { "" }        #It is necessary to change the data output format in the SNPMWalk function
                #"Active Alerts" = if ($pstatus.Length -ne 0) { "$pstatus" } else { "" }    #It is necessary to change the data output format in the SNPMWalk function
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
#endregion

#region Report Generation
    #HTML
    If($HtmlFileReport) {
        $htmlContent = $DataHtmlReport | ConvertTo-Html -Title "Printer Status Report" -Head $Header -PostContent "<p>Report generated: $(Get-Date)</p>"
    
        # Fix HTML encoding
        $htmlContent = $htmlContent -replace '&lt;', '<' -replace '&#39;', "'" -replace '&gt;', '>'
        $filePrefixhtml = Split-Path -Path $selectedFile -Leaf
        $filePrefixhtml = $filePrefixhtml -replace '.txt', '' -replace '.csv', ''
        $filenamehtml = "$ReportDir\Printers_report_$filePrefixhtml $(Get-Date -Format 'yyyyMMddHHmmss').html"
        $htmlContent | Set-Content -Path $filenamehtml -Force -Encoding UTF8
        $DataHtmlReport.Clear()

        if (Test-Path $filenamehtml) {
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
    $end = Get-Date
    $duration = $end - $start
    Write-Log "=== Monitoring completed in $($duration.TotalSeconds) seconds ==="
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

#endregion


