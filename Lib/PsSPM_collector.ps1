function Get-PrinterModelOIDSet {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Model,
        [Parameter(Mandatory=$true)]
        [hashtable]$OIDMapping,
        [Parameter(Mandatory=$true)]
        [hashtable]$ModelPatterns
    )

    foreach ($modelKey in $ModelPatterns.Keys) { if ($Model -match $modelKey) {return $OIDMapping[$ModelPatterns[$modelKey]]} }
    return $OIDMapping["Default"]
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
        $totalValue   = [int]$Total
        $currentValue = [int]$Current

        if ($totalValue -le 0 -or $currentValue -le 0) { return 0 }

        $percentage = [math]::Round(($currentValue / $totalValue) * 100)
        return $percentage
    }
    catch { return $null }
}

function Update-TonerLevels {
    param([object]$PrinterData)

    $TonerLevels = @{}
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
        $totalKey   = $cartridgeConfig[$key][0]
        $currentKey = $cartridgeConfig[$key][1]

        if ($PrinterData.ContainsKey($totalKey) -and $PrinterData.ContainsKey($currentKey)) {
            $TonerLevels[$key] = Get-TonerPercentage -Total $PrinterData[$totalKey] -Current $PrinterData[$currentKey]
        }
    }
    return $TonerLevels
}

function Get-PrinterData {
    param (
        [string]$TargetHost,
        [hashtable]$OidMapping,
        [hashtable]$ModelPatterns,
        [int]$SnmpTimeoutMs,
        [int]$SnmpMaxAttempts,
        [int]$SnmpDelayMs
    )

    $results = [System.Collections.Generic.Dictionary[string, object]]::new()

    try {
        $modelResp = Get-SnmpData -Target $TargetHost -Oid $OidMapping["Default"].Model -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs

        if (-not $modelResp.Success) {
            Write-Log "[$TargetHost] Failed to get printer model" "WARNING"
            return $null
        }

        $model = $modelResp.Result.Data.ToString()
        $results["Model"] = $model
        $results["IPAddress"] = $TargetHost

        $nameResp = Get-SnmpData -Target $TargetHost -Oid $OidMapping["Default"].PName -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs

        if (-not $nameResp.Success) {
            Write-Log "[$TargetHost] Failed to get printer name" "WARNING"
            return $null
        }

        $results["PName"] = $nameResp.Result.Data.ToString()

        $oidSet = Get-PrinterModelOIDSet -Model $model -OIDMapping $OidMapping -ModelPatterns $ModelPatterns
        if (-not $oidSet) {
            Write-Log "[$TargetHost] No OID set for model: $model" "WARNING"
            return $null
        }

        function Resolve-SnmpValue {
            param($Name, $ValueObj)

            if ($Name -in @("Display", "Status")) {
                try {
                    if ($GetBulkWalk) {
                        $bulkResult = Get-SnmpBulkWalkWithEncoding -Target $TargetHost -Oid $oidSet[$Name] -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs
                        return $bulkResult.Value
                    }
                    else {
                        Write-Log "[$TargetHost] Display/Status requires GetBulkWalk=true or regular SNMP GET" "WARNING"
                        return ""
                    }
                }
                catch { 
                    return "Error"
                }
            }

            if (-not $ValueObj.Success) {
                Write-Log "[$TargetHost] SNMP error for $Name" "WARNING"
                return "Error"
            }

            $data = $ValueObj.Result.Data
            if ($null -eq $data) { return "Error" }

            switch -CaseSensitive ($Name) {

                "PMac" {
                    try {
                        $raw = $data.GetRaw()
                        return ([System.BitConverter]::ToString($raw).Replace('-', ':'))
                    }
                    catch { return "Error" }
                }

                "Serial" {
                    try {
                        return ($data.ToString().Replace('??','') -split '-')[0]
                    }
                    catch { return "Error" }
                }

                default {
                    try { return $data.ToString() }
                    catch { return "Error" }
                }
            }
        }

        foreach ($item in $oidSet.GetEnumerator()) {

            $Name = $item.Name
            $Oid  = $item.Value

            if ($Name -in @("Display", "Status")) {
                $results[$Name] = Resolve-SnmpValue -Name $Name -ValueObj $null
                continue
            }

            try {
                $resp = Get-SnmpData -Target $TargetHost -Oid $Oid -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs
                $results[$Name] = Resolve-SnmpValue -Name $Name -ValueObj $resp
            }
            catch {
                Write-Log "[$TargetHost] Exception in Get-SnmpData for $Name : $($_.Exception.Message)" "WARNING"
                $results[$Name] = "Error"
                continue
            }
        }

        return $results
    }
    catch {
        Write-Log "[$TargetHost] Fatal error in Get-PrinterData: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Invoke-PrinterWorker {
    param(
        [System.Collections.Concurrent.ConcurrentQueue[object]]$Queue,
        [System.Collections.Concurrent.ConcurrentBag[object]]$Results,
        [hashtable]$OidMapping,
        [hashtable]$ModelPatterns,
        [int]$SnmpTimeoutMs,
        [int]$SnmpMaxAttempts,
        [int]$SnmpDelayMs,
        [bool]$GetBulkWalk
    )

    while ($true) {
        $ref = [ref]$null
        if (-not $Queue.TryDequeue($ref)) { break }

        $item = $ref.Value
        $ip   = $item.IPAddress
        $idx  = $item.Index

        $printerErrors = [System.Collections.Generic.List[string]]::new()

        function Write-Log {
            param($Message, $Level = "INFO")

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"
            <#switch ($Level) {
                "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
                "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
                #default   { Write-Host $logMessage -ForegroundColor Green }
            }#>

            if ($Level -eq 'ERROR') {
                $printerErrors.Add($logMessage)
            }
        }

        try {
            Write-Log "Starting SNMP for $ip" "INFO" $ip

            $PrinterData = Get-PrinterData -TargetHost $ip -OidMapping $OidMapping -ModelPatterns $ModelPatterns -SnmpTimeoutMs $SnmpTimeoutMs -SnmpMaxAttempts $SnmpMaxAttempts -SnmpDelayMs $SnmpDelayMs

            if ($PrinterData) {
                $tonerLevels = Update-TonerLevels -PrinterData $PrinterData

                $Results.Add([PSCustomObject]@{
                    IPAddress   = $ip
                    TCPStatus   = "Online"
                    PrinterData = $PrinterData
                    TonerLevels = $tonerLevels
                    Error       = $null
                    Index       = $idx
                    AllErrors   = $printerErrors.ToArray()
                })
            }
            else {
                $Results.Add([PSCustomObject]@{
                    IPAddress   = $ip
                    TCPStatus   = "Error"
                    PrinterData = $null
                    TonerLevels = $null
                    Error       = "Failed to get printer data"
                    Index       = $idx
                    AllErrors   = $printerErrors.ToArray()
                })
            }
        }
        catch {
            $Results.Add([PSCustomObject]@{
                IPAddress   = $ip
                TCPStatus   = "Error"
                PrinterData = $null
                TonerLevels = $null
                Error       = $_.Exception.Message
                Index       = $idx
                AllErrors   = $printerErrors.ToArray()
            })
        }
    }
}

function Get-PrintersDataParallel {
    param (
        [Parameter(Mandatory=$true)]
        [array]$CheckPrinters,
        [Parameter(Mandatory=$true)]
        [hashtable]$OidMapping,
        [hashtable]$ModelPatterns,
        [int]$SnmpTimeoutMs = 5000,
        [int]$SnmpMaxAttempts = 3,
        [int]$SnmpDelayMs = 2000,
        [int]$MaxThreads = 20
    )

    $online  = $CheckPrinters | Where-Object { $_.Connected -eq $true }
    $offline = $CheckPrinters | Where-Object { $_.Connected -eq $false }

    $results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $index = 0

    foreach ($p in $offline) {
        $index++
        $results.Add([PSCustomObject]@{
            IPAddress   = $p.Device
            TCPStatus   = "Offline"
            PrinterData = $null
            TonerLevels = $null
            Error       = $null
            Index       = $index
            AllErrors   = @()
        })
    }

    if ($online.Count -eq 0) {
        Write-Log "No online printers to check" "WARNING"
        return $results.ToArray()
    }

    $queue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    foreach ($p in $online) {
        $index++
        $queue.Enqueue([PSCustomObject]@{
            IPAddress = $p.Device
            Index     = $index
        })
    }

    # --- InitialSessionState ---
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    foreach ($fn in @(
        'Get-PrinterModelOIDSet',
        'Get-SnmpData',
        'Get-SnmpBulkWalkWithEncoding',
        'Get-TonerPercentage',
        'Update-TonerLevels',
        'Get-PrinterData',
        'Invoke-PrinterWorker',
        'Write-Log'
    )) {
        $cmd = Get-Command $fn -ErrorAction SilentlyContinue
        if ($cmd) {
            $entry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry $cmd.Name, $cmd.Definition
            $iss.Commands.Add($entry)
        }
    }

    $threads = [Math]::Min($MaxThreads, $online.Count)
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $threads, $iss, $Host)
    $pool.Open()

    $runspaces = New-Object System.Collections.Generic.List[object]

    $null = for ($i = 1; $i -le $threads; $i++) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool

        $ps.AddCommand("Invoke-PrinterWorker")
        $ps.AddParameter("Queue", $queue)
        $ps.AddParameter("Results", $results)
        $ps.AddParameter("OidMapping", $OidMapping)
        $ps.AddParameter("ModelPatterns", $ModelPatterns)
        $ps.AddParameter("SnmpTimeoutMs", $SnmpTimeoutMs)
        $ps.AddParameter("SnmpMaxAttempts", $SnmpMaxAttempts)
        $ps.AddParameter("SnmpDelayMs", $SnmpDelayMs)
        $ps.AddParameter("GetBulkWalk", $GetBulkWalk)

        $runspaces.Add([PSCustomObject]@{
            PS     = $ps
            Handle = $ps.BeginInvoke()
        })
    }

    Write-Log "Started $threads workers for $($online.Count) printers"
    $total = $online.Count

    $lastCompleted = -1

    while ($runspaces.Handle.IsCompleted -contains $false) {
        $currentCompleted = ($results | Where-Object { $_.TCPStatus -eq "Online" }).Count
        
        # Выводим прогресс только если изменился
        if ($currentCompleted -ne $lastCompleted) {
            $progress = [math]::Round(($currentCompleted / $total) * 100, 2)
            Write-Log "Processed: $progress% ($currentCompleted / $total)"
            $lastCompleted = $currentCompleted
        }
        
        Start-Sleep -Milliseconds 500
    }

    # Финальный подсчёт уже не нужен, используем последнее значение
    $completed = ($results | Where-Object { $_.TCPStatus -eq "Online" }).Count
    $progress = [math]::Round(($completed / $total) * 100, 2)
    Write-Log "FINAL: $progress% ($completed / $total)"

    if ($completed -ne $total) {
        Write-Log "WARNING: Only $completed out of $total hosts are online"
    }

    foreach ($r in $runspaces) {
        try { $null = $r.PS.EndInvoke($r.Handle) } catch {}
        finally { $r.PS.Dispose() }
    }

    $pool.Close()
    $pool.Dispose()

    return $results.ToArray()
}