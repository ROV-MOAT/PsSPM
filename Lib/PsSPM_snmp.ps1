<#
.SYNOPSIS
    PowerShell SNMP Printer Monitoring and Reporting Script / PsSPM (ROV-MOAT)

.LICENSE
    Distributed under the MIT License. See the accompanying LICENSE file or https://github.com/ROV-MOAT/PsSPM/blob/main/LICENSE

.DESCRIPTION
    SNMP Get and BulkWal functions
#>

function Get-SnmpData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [string]$Community = "public",
        [Parameter(Mandatory=$true)]
        [string]$Oid,
        [int]$UDPport = 161,
        [int]$SnmpTimeoutMs = 5000,
        [int]$SnmpMaxAttempts = 3,
        [int]$SnmpDelayMs = 2000
    )

    try {
        $IP = [System.Net.IPAddress]::Parse($Target)
    }
    catch {
        $errorMsg = "Invalid IP address: $Target - $($_.Exception.Message)"
        Write-Log $errorMsg -Level "ERROR"
        return [PSCustomObject]@{
            Success      = $false
            Result       = $null
            ErrorMessage = $errorMsg
        }
    }

    $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)
    $oidObj   = [Lextm.SharpSnmpLib.ObjectIdentifier]::new($Oid)

    for ($attempt = 1; $attempt -le $SnmpMaxAttempts; $attempt++) {

        Write-Log "SNMP query attempt $attempt of $SnmpMaxAttempts for $Target (OID: $Oid)" -Level "INFO"

        try {
            $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
            $vList.Add([Lextm.SharpSnmpLib.Variable]::new($oidObj)) | Out-Null

            $result = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
                [Lextm.SharpSnmpLib.VersionCode]::V2,
                $endpoint,
                [Lextm.SharpSnmpLib.OctetString]::new($Community),
                $vList,
                $SnmpTimeoutMs
            )

            Write-Log "SNMP query successful for $Target (OID: $Oid) on attempt $attempt" -Level "INFO"

            return [PSCustomObject]@{
                Success      = $true
                Result       = $result
                ErrorMessage = $null
            }
        }
        catch {
            $lastError = $_.Exception.Message
            Write-Log "SNMP query failed for $Target on attempt $attempt : $lastError" -Level "WARNING"

            if ($attempt -lt $SnmpMaxAttempts) {
                Write-Log "Retrying in $SnmpDelayMs ms..." -Level "WARNING"
                Start-Sleep -Milliseconds $SnmpDelayMs
            }
        }
    }

    $finalErrorMsg = "SNMP query failed for $Target (OID: $Oid) after $SnmpMaxAttempts attempts. Last error: $lastError"
    Write-Log $finalErrorMsg -Level "ERROR"

    return [PSCustomObject]@{
        Success      = $false
        Result       = $null
        ErrorMessage = $finalErrorMsg
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
        [int]$SnmpTimeoutMs = 5000,
        [int]$SnmpMaxAttempts = 3,
        [int]$SnmpDelayMs = 2000
    )

    $encodings = @(
        [System.Text.Encoding]::UTF8,
        [System.Text.Encoding]::ASCII,
        [System.Text.Encoding]::GetEncoding(1251)
    )

    try {
        $IP = [System.Net.IPAddress]::Parse($Target)
    }
    catch {
        $errorMsg = "Invalid IP address: $Target - $($_.Exception.Message)"
        Write-Log $errorMsg -Level "ERROR"
        return @()
    }

    $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)

    for ($attempt = 1; $attempt -le $SnmpMaxAttempts; $attempt++) {

        Write-Log "SNMP BulkWalk attempt $attempt of $SnmpMaxAttempts for $Target (OID: $Oid)" -Level "INFO"

        try {
            $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()

            $null = [Lextm.SharpSnmpLib.Messaging.Messenger]::BulkWalk(
                [Lextm.SharpSnmpLib.VersionCode]::V2,
                $endpoint,
                [Lextm.SharpSnmpLib.OctetString]::new($Community),
                [Lextm.SharpSnmpLib.OctetString]::Empty,
                [Lextm.SharpSnmpLib.ObjectIdentifier]::new($Oid),
                $vList,
                $SnmpTimeoutMs,
                $MaxRepetitions,
                [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree,
                $null,
                $null
            )

            Write-Log "SNMP BulkWalk successful on attempt $attempt, items: $($vList.Count)" -Level "INFO"

            $output = foreach ($item in $vList) {
                $data  = $item.Data
                $value = $null

                if ($data.TypeCode -eq [Lextm.SharpSnmpLib.SnmpType]::OctetString) {
                    foreach ($enc in $encodings) {
                        try {
                            $value = $data.ToString($enc)
                            break
                        }
                        catch {}
                    }

                    if (-not $value) {
                        $value = $data.ToString()
                    }
                }
                else {
                    $value = $data.ToString()
                }

                if ($value.Trim() -ne "") {
                    [PSCustomObject]@{
                        Success = $true
                        Oid   = $item.Id.ToString()
                        Value = $value
                    }
                }
            }

            return $output
        }
        catch {
            $lastError = $_.Exception.Message
            Write-Log "SNMP BulkWalk failed for $Target on attempt $attempt : $lastError" -Level "WARNING"

            if ($attempt -lt $SnmpMaxAttempts) {
                Write-Log "Retrying in $SnmpDelayMs ms..." -Level "WARNING"
                Start-Sleep -Milliseconds $SnmpDelayMs
            }
        }
    }

    $finalErrorMsg = "SNMP BulkWalk failed for $Target (OID: $Oid) after $SnmpMaxAttempts attempts. Last error: $lastError"
    Write-Log $finalErrorMsg -Level "ERROR"
    return @()
}