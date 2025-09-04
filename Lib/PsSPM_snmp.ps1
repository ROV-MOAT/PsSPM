<#
.SYNOPSIS
    PsSPM(ROV-MOAT) - PowerShell SNMP Printer Monitoring and Reporting Script

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

    $attempt = 0
    
    while ($attempt -lt $SnmpMaxAttempts) {
        $attempt++
        try {
            # Set up SNMP manager
            $IP = [System.Net.IPAddress]::Parse($Target)
            $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)
            $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
            $null = $vList.Add([Lextm.SharpSnmpLib.Variable]::new([Lextm.SharpSnmpLib.ObjectIdentifier]::new($Oid)))

            $result = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
                [Lextm.SharpSnmpLib.VersionCode]::V2,
                $endpoint,
                [Lextm.SharpSnmpLib.OctetString]::new($Community),
                $vList,
                $SnmpTimeoutMs
            )
            return [PSCustomObject]@{
                Success = $true
                result = $result.Data.ToString()
            }
        }
        catch {
            Write-Log "SNMP query failed for $Target (OID: $Oid): $_" -Level "ERROR"
            if ($attempt -lt $SnmpMaxAttempts) { Write-Log "SNMP Get - repeat in $SnmpDelayMs milliseconds..." -Level "WARNING"; Start-Sleep -Milliseconds $SnmpDelayMs }
        }
    }
    return [PSCustomObject]@{
        Success = $false
        result = $null
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
        [int]$SnmpTimeoutMs = 5000
    )

    $result_out = [System.Collections.Generic.List[PSObject]]::new()
    $encodings = @([System.Text.Encoding]::UTF8, [System.Text.Encoding]::ASCII, [System.Text.Encoding]::GetEncoding(1251))

    try {
        # Set up SNMP manager
        $IP = [System.Net.IPAddress]::Parse($Target)
        $endpoint = [System.Net.IPEndPoint]::new($IP, $UDPport)
        $vList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()

        # Perform the walk
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
        foreach ($result in $vList) {
            $data = $result.Data
            $value = $null

            # Handle OctetString encoding
            if ($data.TypeCode -eq [Lextm.SharpSnmpLib.SnmpType]::OctetString) {
                foreach ($enc in $encodings) {
                    try { $value = $data.ToString($enc); break }
                    catch { $value = $data.ToString() }
                }
                $null = $result_out.Add($value)
            }
        }
        $cleanArray = $result_out.Where( {$_.Trim() -ne ""} )
        return $cleanArray
    }
    catch { Write-Log "SNMP Walk query failed for $Target (OID: $Oid): $_" -Level "ERROR"; return "Error" }
}