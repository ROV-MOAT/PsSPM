<#
.SYNOPSIS
    PowerShell SNMP Printer Monitoring and Reporting Script / PsSPM (ROV-MOAT)

.LICENSE
    Distributed under the MIT License. See the accompanying LICENSE file or https://github.com/ROV-MOAT/PsSPM/blob/main/LICENSE

.DESCRIPTION
    Report functions
#>

function EncodeHtml {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) { return $null }

    $s = $Value.ToString()
    if ([string]::IsNullOrEmpty($s)) { return $s }

    # System.Web в Windows PowerShell
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    return [System.Web.HttpUtility]::HtmlEncode($s)
}

function Format-Value {
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Value,
        [string]$Default = "Error"
    )

    if ($null -eq $Value -or $Value -eq "") { return $null }

    if ($Value -like 'Error') {
        return "<span class='error'>$((EncodeHtml $Default))</span>"
    }

    $encoded = EncodeHtml $Value
    return "$encoded"
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

    $encoded = EncodeHtml $tonerValue

    switch ($tonerValue) {
        { $_ -gt 49 }                 { return "<span class='toner-high'>$encoded%</span>" }
        { $_ -gt 10 -and $_ -le 49 }  { return "<span class='toner-medium'>$encoded%</span>" }
        { $_ -ge 0 -and $_ -le 10 }   { return "<span class='toner-low'>$encoded%</span>" }
        default                       { return $null }
    }
}

function Format-Status {
    param([string]$TcpStatus)

    $encoded = EncodeHtml $TcpStatus

    $class = switch ($TcpStatus) {
        "Online"  { "online" }
        "Offline" { "offline" }
        default   { "error" }
    }

    return "<span class='$class'>$encoded</span>"
}

function Add-HtmlPrinterRowString {
    param(
        [System.Collections.Generic.List[string]]$DataHtmlReport,
        [object]$Collector
    )

    function _li {
        param($arr)
        if (-not $arr -or -not $arr.Count) { return "" }
        return (($arr | ForEach-Object { "<li>$_</li>" }) -join "")
    }

    $PrinterIP = $Collector.IPAddress
    $TcpStatus = $Collector.TCPStatus
    $TonerLevels = $Collector.TonerLevels
    $htmlDisplay = _li $Collector.PrinterData.Display
    $htmlStatus  = _li $Collector.PrinterData.Status
    $htmlError   = _li $Collector.SnmpErrors

    $displayCell = if ($htmlDisplay) {
        "<a class='show-link' href='http://$PrinterIP' target='_blank' data-message='<ul>$htmlDisplay</ul>'>Show</a>"
    } else { "" }

    $statusCell = if ($htmlStatus) {
        "<a class='show-link' href='http://$PrinterIP' target='_blank' data-message='<ul>$htmlStatus</ul>'>Show</a>"
    } else { "" }

    $errorCell = if ($htmlError) {
        "<a class='show-link' href='#' target='_blank' data-message='<ul>$htmlError</ul>'>E</a>"
    } else { "" }

    $cToner = if ($TonerLevels.TC) { "<span class='container'>$(Format-TonerLevel $TonerLevels.TC)</span>" } else { "" }
    $mToner = if ($TonerLevels.TM) { "<span class='container'>$(Format-TonerLevel $TonerLevels.TM)</span>" } else { "" }
    $yToner = if ($TonerLevels.TY) { "<span class='container'>$(Format-TonerLevel $TonerLevels.TY)</span>" } else { "" }
    $kToner = if ($TonerLevels.TK) { "<span class='container'>$(Format-TonerLevel $TonerLevels.TK)</span>" } else { "" }

    $drum = if ($TonerLevels.DC -or $TonerLevels.DM -or $TonerLevels.DY -or $TonerLevels.DK -or $TonerLevels.DKU) {
        "<span class='container'>$(Format-TonerLevel $TonerLevels.DC) $(Format-TonerLevel $TonerLevels.DM) $(Format-TonerLevel $TonerLevels.DY) $(Format-TonerLevel $TonerLevels.DKU) $(Format-TonerLevel $TonerLevels.DK)</span>"
    } else { "" }

    $row = @"
        <tr>
            <td><a class='printer-link' href='http://$PrinterIP' target='_blank'>$PrinterIP</a></td>
            <td>$(Format-Status -TcpStatus $TcpStatus)</td>
            <td><a class='printer-link' href='http://$PrinterIP' target='_blank'>$($Collector.PrinterData.PName)</a></td>
            <td><span>$(Format-Value $Collector.PrinterData.PMac)</span></td>
            <td>$(Format-Value $Collector.PrinterData.Model)</td>
            <td><span>$(Format-Value $Collector.PrinterData.Serial)</span></td>
            <td>$(Format-Value $Collector.PrinterData.BlackCount)</td>
            <td>$(Format-Value $Collector.PrinterData.ColorCount)</td>
            <td>$(Format-Value $Collector.PrinterData.TotalCount)</td>
            <td>$cToner</td>
            <td>$mToner</td>
            <td>$yToner</td>
            <td>$kToner</td>
            <td>$drum</td>
            <td>$displayCell</td>
            <td>$statusCell</td>
            <td>$errorCell</td>
        </tr>
"@

    $DataHtmlReport.Add($row) | Out-Null
}

function Add-CsvPrinterRow {
    param(
        [System.Collections.Generic.List[PSObject]]$DataCsvReport,
        [string]$PrinterIP,
        [string]$TcpStatus,
        [object]$PrinterData,
        [object]$TonerLevels,
        [array] $PrinterErrors,
        [string]$CsvPath,
        [int]$CsvBufferSize = 200
    )

    $row = [PSCustomObject]@{
        "IP"        = $PrinterIP
        "Status"    = $TcpStatus
        "Name"      = $PrinterData.PName
        "MAC"       = $PrinterData.PMac
        "Model"     = $PrinterData.Model
        "S/N"       = $PrinterData.Serial
        "Black"     = $PrinterData.BlackCount
        "Color"     = $PrinterData.ColorCount
        "Total"     = $PrinterData.TotalCount
        "C Toner %" = $TonerLevels.TC
        "M Toner %" = $TonerLevels.TM
        "Y Toner %" = $TonerLevels.TY
        "K Toner %" = $TonerLevels.TK
        "CMYK DrumKit %" = "$($TonerLevels.DC) $($TonerLevels.DM) $($TonerLevels.DY) $($TonerLevels.DKU) $($TonerLevels.DK)"
        "Display" = if ($PrinterData.Display.Count) { "$($PrinterData.Display)" } else { "" }
        "Active Alerts" = if ($PrinterData.Status.Count) { "$($PrinterData.Status)" } else { "" }
        "Collector Error" = "$PrinterErrors"
    }

    $DataCsvReport.Add($row) | Out-Null

    if ($DataCsvReport.Count -ge $CsvBufferSize) {
        $DataCsvReport | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
        $DataCsvReport.Clear()
    }
}
