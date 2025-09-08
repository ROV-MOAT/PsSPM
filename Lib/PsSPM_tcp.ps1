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

function Test-TcpConnectionParallel {
    param(
        [array]$Devices = @(),
        [int]$Port = 80,
        [int]$TcpTimeoutMs = 500,
        [int]$MaxRetries = 3,
        [int]$RetryDelayMs = 3000,
        [int]$Threads = 5
    )
    $start = Get-Date
    # Create RunspacePool
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
    $runspacePool.Open()
    $runspaces = @()

    # ScriptBlock for connection testing
    $scriptBlock = {
        param($Device, $Port, $TcpTimeoutMs, $MaxRetries, $RetryDelayMs)

        function Connect-TcpClient {
            param($HostName, $Port, $TcpTimeoutMs, $MaxRetries, $RetryDelayMs)

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
                    $completedTask = [System.Threading.Tasks.Task]::WaitAny($connectTask, [System.Threading.Tasks.Task]::Delay($TcpTimeoutMs))
                    
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
                        throw [System.TimeoutException]::new("Connection attempt timed out after $TcpTimeoutMs ms")
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
        Connect-TcpClient -HostName $Device -Port $Port -TcpTimeoutMs $TcpTimeoutMs -MaxRetries $MaxRetries -RetryDelayMs $RetryDelayMs
    }

    # Create and start runspaces
    foreach ($Device in $Devices) {
        $powershell = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($Device).AddArgument($Port).AddArgument($TcpTimeoutMs).AddArgument($MaxRetries).AddArgument($RetryDelayMs)
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