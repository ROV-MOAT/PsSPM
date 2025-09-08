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
        [int]$SnmpDelayMs = 2000,
        [switch]$MacAdr
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
            if ($MacAdr) {
                $Value = $result.data.GetRaw()
                $rawBytes = [System.BitConverter]::ToString($Value).Replace('-', '')
                #$dectohex = $rawBytes.ToString()
                $hex = ($rawBytes.ToString() -replace '(.{2})', '$1:') -replace ':$'

                return [PSCustomObject]@{
                    Success = $true
                    result = $hex
                }
            
            } else {
                return [PSCustomObject]@{
                    Success = $true
                    result = $result.Data.ToString().Replace('??', '')
                }
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
    catch { Write-Log "SNMP Walk query failed for $Target (OID: $Oid): $_" -Level "ERROR"; return "<span class='error'>Error</span>" }
}