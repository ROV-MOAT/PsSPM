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
                    Message   = "Success"
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
        finally {
            if ($client) {
                try { $client.Dispose() } catch {}
            }
        }
    }
    
    $finalErrorMsg = "TCP query failed for $HostName after $attempt attempts. Last error: $lastException"
    Write-Log $finalErrorMsg -Level "ERROR"

    return [PSCustomObject]@{
        Device    = $HostName
        Port      = $Port
        Connected = $false
        Message   = $lastException.Message
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

    # --- Create queue and result bag ---
    $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    foreach ($d in $Devices) { $queue.Enqueue($d) }

    $results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

    # --- Prepare InitialSessionState with external function ---
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    $func = Get-Command Connect-TcpClient
    $funcEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry $func.Name, $func.Definition
    $iss.Commands.Add($funcEntry)

    # --- Create runspace pool ---
    $minThreads = [Math]::Min($Threads, $Devices.Count)
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $minThreads, $iss, $Host)
    $runspacePool.Open()

    $runspaces = New-Object System.Collections.Generic.List[object]

    # --- Worker script ---
    $scriptBlock = {
        param($queue, $results, $Port, $TcpTimeoutMs, $MaxRetries, $RetryDelayMs)

        function Write-Log {
            param($Message, $Level = "INFO")

            <#$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"
            switch ($Level) {
                "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
                #"WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
                #default   { Write-Host $logMessage -ForegroundColor Green }
            }#>
        }

        while ($true) {
            $ref = [ref] ""
            if (-not $queue.TryDequeue($ref)) { break }

            $device = $ref.Value
            $res = Connect-TcpClient -HostName $device -Port $Port -TcpTimeoutMs $TcpTimeoutMs -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs
            $results.Add($res)
        }
    }

    # --- Start workers ---
    $null = for ($i = 1; $i -le $minThreads; $i++) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $runspacePool
        $ps.AddScript($scriptBlock).AddArgument($queue).AddArgument($results).AddArgument($Port).AddArgument($TcpTimeoutMs).AddArgument($MaxRetries).AddArgument($RetryDelayMs)

        $runspaces.Add([PSCustomObject]@{
            PowerShell = $ps
            Handle = $ps.BeginInvoke()
        })
    }

    Write-Log "Started $minThreads workers for $($Devices.Count) devices"
    $total = $Devices.Count
    $lastCompleted = -1
    
    while ($runspaces.Handle.IsCompleted -contains $false) {
        $currentCompleted = $results.Count

        # Выводим прогресс только если изменился
        if ($currentCompleted -ne $lastCompleted) {
            $progress = [math]::Round(($currentCompleted / $total) * 100, 2)
            Write-Log "Processed: $progress% ($currentCompleted / $total)"
            $lastCompleted = $currentCompleted
        }
        
        Start-Sleep -Milliseconds 500
    }

    $final = $results.Count
    $finalProgress = [math]::Round(($final / $total) * 100, 2)
    Write-Log "FINAL: $finalProgress% ($final / $total)"

    $finalresults = ($results | Where-Object { $_.Connected -eq $true }).Count
    if ($finalresults -ne $total) {
        Write-Log "WARNING: Only $finalresults out of $total hosts are online" -level "WARNING"
    }

    # --- Finalize ---
    foreach ($r in $runspaces) {
        try { $null = $r.PowerShell.EndInvoke($r.Handle) } catch {}
        finally { $r.PowerShell.Dispose() }
    }

    $runspacePool.Close()
    $runspacePool.Dispose()

    return $results.ToArray()
}