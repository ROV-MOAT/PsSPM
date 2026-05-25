<#
.SYNOPSIS
    PowerShell SNMP Printer Monitoring and Reporting Script / PsSPM (ROV-MOAT)

.LICENSE
    Distributed under the MIT License. See the accompanying LICENSE file or https://github.com/ROV-MOAT/PsSPM/blob/main/LICENSE

.DESCRIPTION
    TCP test function

    Usage:
        Test-TcpConnectionParallel -Devices "10.10.10.10, 11.11.11.11" -Port 80 -TcpTimeoutMs 500 -MaxRetries 3 -RetryDelayMs 3000 -Threads 10
#>
function Connect-TcpClient {
    param(
        [string]$HostName,
        [int]$Port,
        [int]$TcpTimeoutMs,
        [int]$MaxRetries,
        [int]$RetryDelayMs
    )

    $attempt = 0
    $lastException = $null

    while ($attempt -lt $MaxRetries) {
        $attempt++
        $client = $null

        try {
            $client = [System.Net.Sockets.TcpClient]::new()
            $task = $client.ConnectAsync($HostName, $Port)

            $completed = [System.Threading.Tasks.Task]::WaitAny($task, [System.Threading.Tasks.Task]::Delay($TcpTimeoutMs))

            if ($completed -eq 0) {
                if ($task.IsFaulted) { throw $task.Exception.InnerException }

                return [PSCustomObject]@{
                    Device    = $HostName
                    Port      = $Port
                    Connected = $true
                    Exception = $null
                }
            }
            else {
                write-log "$HostName Timeout after: $TcpTimeoutMs ms / Attempt: $attempt" -Level "WARNING"
                throw [System.TimeoutException]::new("Timeout after: $TcpTimeoutMs ms / Attempt: $attempt")
            }
        }
        catch {
            $lastException = $_
            if ($attempt -lt $MaxRetries) { Start-Sleep -Milliseconds $RetryDelayMs }
        }
        finally { if ($client) { try { $client.Dispose() } catch {} } }
    }
    
    $finalErrorMsg = "TCP query failed for $HostName after $attempt attempts. Last error: $lastException"
    Write-Log $finalErrorMsg -Level "ERROR"

    return [PSCustomObject]@{
        Device    = $HostName
        Port      = $Port
        Connected = $false
        Exception = $lastException
    }
}

function Test-TcpConnectionParallel {
    param(
        $Devices = @(),
        [int]$Port = 80,
        [int]$TcpTimeoutMs = 500,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 3000,
        [int]$Threads = 50
    )

    if ($Devices -is [string]) { $Devices = $Devices -split '\s*,\s*' | Where-Object { $_ } }

    if ($Devices.Count -eq 0) {
        Write-Log "No devices to test"
        return @()
    }

    # --- Shared collections ---
    $queue   = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

    foreach ($d in $Devices) { $queue.Enqueue($d) }

    # --- Shared parameters ---
    $shared = [PSCustomObject]@{
        Port         = $Port
        TcpTimeoutMs = $TcpTimeoutMs
        MaxRetries   = $MaxRetries
        RetryDelayMs = $RetryDelayMs
    }

    # --- InitialSessionState ---
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    # Import Connect-TcpClient
    $func = Get-Command Connect-TcpClient
    $funcEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry $func.Name, $func.Definition
    $iss.Commands.Add($funcEntry)

    # Worker function — минимальная логика
    $worker = {
        param($queue, $results, $shared)

        function Write-Log {
            param($Message, $Level = "INFO")

            <#$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = '{0} [{1}] - {2}' -f $timestamp, $Level, $Message
            switch ($Level) {
                "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
                "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
                default   { Write-Host $logMessage -ForegroundColor Green }
            }#>
        }

        while ($true) {
            $ref = [ref]::new("")
            if (-not $queue.TryDequeue($ref)) { break }

            $dev = $ref.Value
            $res = Connect-TcpClient -HostName $dev -Port $shared.Port -TcpTimeoutMs $shared.TcpTimeoutMs -MaxRetries $shared.MaxRetries -RetryDelayMs $shared.RetryDelayMs
            $results.Add($res)
        }
    }

    $workerEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry "Worker", $worker.ToString()
    $iss.Commands.Add($workerEntry)

    # --- Runspace pool ---
    $max = [Math]::Min($Threads, $Devices.Count)
    $pool = [runspacefactory]::CreateRunspacePool(1, $max, $iss, $Host)
    $pool.Open()

    $runspaces = New-Object System.Collections.Generic.List[object]

    # --- Start workers ---
    $null = for ($i = 1; $i -le $max; $i++) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        $ps.AddCommand("Worker").AddArgument($queue).AddArgument($results).AddArgument($shared)

        $runspaces.Add([PSCustomObject]@{
            PS     = $ps
            Handle = $ps.BeginInvoke()
        })
    }

    # --- Progress loop ---
    $total = $Devices.Count
    $last  = -1
    $spin  = [System.Threading.SpinWait]::new()

    while ($runspaces.Handle.IsCompleted -contains $false) {
        $cur = $results.Count
        if ($cur -ne $last) {
            $pct = [math]::Round(($cur / $total) * 100, 2)
            Write-Log "Processed: $pct% ($cur / $total)"
            $last = $cur
        }
        $spin.SpinOnce()
    }

    $final = $results.Count
    $finalProgress = [math]::Round(($final / $total) * 100, 2)
    Write-Log "FINAL: $finalProgress% ($final / $total)"

    $finalresults = ($results | Where-Object { $_.Connected -eq $true }).Count
    if ($finalresults -ne $total) {
        Write-Log "Only $finalresults out of $total hosts are successful" -level "WARNING"
    }

    # --- Finalize ---
    foreach ($r in $runspaces) {
        try { $null = $r.PS.EndInvoke($r.Handle) } catch {}
        finally { $r.PS.Dispose() }
    }

    $pool.Close()
    $pool.Dispose()

    return $results.ToArray()
}
