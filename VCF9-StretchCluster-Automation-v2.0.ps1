<#
.SYNOPSIS
 This utility is a WPF-based PowerShell 7 tool that helps you **generate, validate, and execute** the JSON payloads required to stretch a VCF 9 management cluster across AZ1 and AZ2. It reads from **SDDC Manager (source of truth)**, optionally connects to **vCenter**, and can also **populate an Excel data-collection workbook**.

.DESCRIPTION
  Fully rolled standalone script. No patcher/hotfix required.

  Includes:
    - SDDC Manager connection
    - vCenter verification and vDS/vmnic detection fallback
    - NSX Manager connection
    - Local AZ2 TEP CIDR/IP/VLAN validation
    - NSX IP pool overlap validation
    - NSX uplink profile validation
    - Warning if generated uplink profile already exists
    - Blocking validation if requested active/standby uplinks are not known from NSX uplink profiles
    - Generate / Validate / Execute workflow

.NOTES
  Requires PowerShell 7+ on Windows for WPF.
  vCenter detection requires VMware.PowerCLI.
#>

[CmdletBinding()]
param([switch]$NoRelaunch)

$Global:VCFStretchVersion = '1.3.9c'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

try {
    $pwsh = (Get-Process -Id $PID -ErrorAction SilentlyContinue).Path
    if (-not $pwsh) { $pwsh = 'pwsh.exe' }
} catch { $pwsh = 'pwsh.exe' }

try {
    if (-not $NoRelaunch -and [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        & $pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "$PSCommandPath" -NoRelaunch
        exit $LASTEXITCODE
    }
} catch {}

# =============================================================================
# State
# =============================================================================
$script:ReportsBase = (Get-Location).Path
$script:RunDir = $null
$Global:LogFile = $null
$script:SddcSession = $null
$script:NsxSession = $null
$script:VCenterVerified = $false
$script:ClusterMap = @{}
$script:DetectedNetwork = $null

# =============================================================================
# Logging / status
# =============================================================================
function New-RunDir {
    param([string]$Base)
    if ([string]::IsNullOrWhiteSpace($Base) -or -not (Test-Path -LiteralPath $Base)) { $Base = (Get-Location).Path }
    $d = Join-Path $Base ("VCFStretch-Run-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Force -Path $d | Out-Null
    $Global:LogFile = Join-Path $d ("VCFStretch-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.log')
    '' | Set-Content -Path $Global:LogFile -Encoding UTF8
    $script:RunDir = $d
    return $d
}

function Write-Log {
    param([Parameter(Mandatory)][string]$Message,[ValidateSet('INFO','WARN','ERROR')][string]$Level='INFO')
    $line = "[{0}][{1}] {2}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    try { if ($Global:LogFile) { Add-Content -Path $Global:LogFile -Value $line -Encoding UTF8 } } catch {}
    try {
        if ($script:txtLog) {
            if ($script:txtLog.Dispatcher.CheckAccess()) {
                $script:txtLog.AppendText("$line`r`n")
                $script:txtLog.ScrollToEnd()
            } else {
                $script:txtLog.Dispatcher.Invoke([Action]{
                    $script:txtLog.AppendText("$line`r`n")
                    $script:txtLog.ScrollToEnd()
                }) | Out-Null
            }
        }
    } catch {}
    Write-Host $line
}

function Set-StatusText {
    param($Label,[string]$Text,[string]$State)
    if (-not $Label) { return }
    $Label.Text = $Text
    if ($State -eq 'OK') { $Label.Foreground = [Windows.Media.Brushes]::LightGreen }
    elseif ($State -eq 'FAIL') { $Label.Foreground = [Windows.Media.Brushes]::Tomato }
    elseif ($State -eq 'WARN') { $Label.Foreground = [Windows.Media.Brushes]::Gold }
    else { $Label.Foreground = [Windows.Media.Brushes]::White }
}

function Has-Module {
    param([string]$Name)
    return [bool](Get-Module -ListAvailable -Name $Name | Select-Object -First 1)
}

function Prereq-Check {
    $isPS7 = $PSVersionTable.PSVersion.Major -ge 7
    if ($isPS7) { Set-StatusText $script:lblPS "PowerShell $($PSVersionTable.PSVersion)" 'OK' } else { Set-StatusText $script:lblPS "PowerShell $($PSVersionTable.PSVersion)" 'FAIL' }
    Set-StatusText $script:lblWPF '.NET/WPF: OK' 'OK'
    if (Has-Module 'ImportExcel') { Set-StatusText $script:lblImpExcel 'ImportExcel: Found' 'OK' } else { Set-StatusText $script:lblImpExcel 'ImportExcel: Not found' 'WARN' }
    if (Has-Module 'VMware.VimAutomation.Core') { Set-StatusText $script:lblPCLI 'VMware.PowerCLI: Found' 'OK' } else { Set-StatusText $script:lblPCLI 'VMware.PowerCLI: Not found' 'WARN' }
    if (Has-Module 'VCF.PowerCLI') { Set-StatusText $script:lblVCFPCLI 'VCF.PowerCLI: Found' 'OK' } else { Set-StatusText $script:lblVCFPCLI 'VCF.PowerCLI: Not found' 'WARN' }
    return $isPS7
}

# =============================================================================
# IP helpers
# =============================================================================
function Test-IPv4String {
    param([string]$Ip)
    if ([string]::IsNullOrWhiteSpace($Ip)) { return $false }
    $parts = $Ip.Trim().Split('.')
    if ($parts.Count -ne 4) { return $false }
    foreach ($p in $parts) {
        $n = 0
        if (-not [int]::TryParse($p, [ref]$n)) { return $false }
        if ($n -lt 0 -or $n -gt 255) { return $false }
    }
    return $true
}

function Convert-IPv4ToUInt32 {
    param([string]$Ip)
    if (-not (Test-IPv4String $Ip)) { throw "Invalid IPv4 address: $Ip" }
    $p = $Ip.Trim().Split('.') | ForEach-Object { [uint32]$_ }
    return (($p[0] -shl 24) -bor ($p[1] -shl 16) -bor ($p[2] -shl 8) -bor $p[3])
}

function Convert-UInt32ToIPv4 {
    param([uint32]$Value)
    return "$(($Value -shr 24) -band 255).$((($Value -shr 16) -band 255)).$((($Value -shr 8) -band 255)).$(($Value -band 255))"
}

function Test-CidrString {
    param([string]$Cidr)
    if ([string]::IsNullOrWhiteSpace($Cidr)) { return $false }
    if ($Cidr -notmatch '^(.+)/(\d{1,2})$') { return $false }
    if (-not (Test-IPv4String $Matches[1])) { return $false }
    $prefix = [int]$Matches[2]
    return ($prefix -ge 1 -and $prefix -le 32)
}

function Get-CidrBounds {
    param([string]$Cidr)
    if (-not (Test-CidrString $Cidr)) { throw "Invalid CIDR: $Cidr" }
    $null = $Cidr -match '^(.+)/(\d{1,2})$'
    $ip = $Matches[1]
    $prefix = [int]$Matches[2]
    $mask = [uint32]([uint32]::MaxValue -shl (32 - $prefix))
    $network = (Convert-IPv4ToUInt32 $ip) -band $mask
    $broadcast = [uint32]($network -bor (-bnot $mask))
    return [pscustomobject]@{
        Network = [uint32]$network
        Broadcast = [uint32]$broadcast
        First = [uint32]($network + 1)
        Last = [uint32]($broadcast - 1)
    }
}

function Test-IPInCidr {
    param([string]$Ip,[string]$Cidr)
    $b = Get-CidrBounds $Cidr
    $i = Convert-IPv4ToUInt32 $Ip
    return ($i -ge $b.Network -and $i -le $b.Broadcast)
}

function Test-RangeOverlap {
    param([uint32]$AStart,[uint32]$AEnd,[uint32]$BStart,[uint32]$BEnd)
    return ($AStart -le $BEnd -and $BStart -le $AEnd)
}

function Find-AvailableRange {
    param([string]$Cidr,[uint32]$Needed,[object[]]$Used,[string]$Gateway)
    $b = Get-CidrBounds $Cidr
    $gwInt = Convert-IPv4ToUInt32 $Gateway
    $cur = $b.First
    while (($cur + $Needed - 1) -le $b.Last) {
        $s = [uint32]$cur
        $e = [uint32]($cur + $Needed - 1)
        if ($gwInt -ge $s -and $gwInt -le $e) { $cur = [uint32]($gwInt + 1); continue }
        $bad = $false
        foreach ($r in $Used) {
            if (Test-RangeOverlap $s $e $r.StartInt $r.EndInt) {
                $cur = [uint32]($r.EndInt + 1)
                $bad = $true
                break
            }
        }
        if (-not $bad) { return [pscustomobject]@{ Start=(Convert-UInt32ToIPv4 $s); End=(Convert-UInt32ToIPv4 $e) } }
    }
    return $null
}

function Assert-Az2PoolFields {
    $pool = ([string]$script:txtTepPoolName.Text).Trim()
    $cidr = ([string]$script:txtTepCidr.Text).Trim()
    $gw = ([string]$script:txtTepGateway.Text).Trim()
    $start = ([string]$script:txtTepRangeStart.Text).Trim()
    $end = ([string]$script:txtTepRangeEnd.Text).Trim()
    $vlanText = ([string]$script:txtTransportVlan.Text).Trim()

    if (@($pool,$cidr,$gw,$start,$end,$vlanText) | Where-Object { [string]::IsNullOrWhiteSpace($_) }) { throw 'AZ2 NSX TEP Pool fields are required.' }
    if (-not (Test-CidrString $cidr)) { throw "Invalid TEP CIDR: $cidr" }
    foreach ($pair in @(@('Gateway',$gw),@('Range Start',$start),@('Range End',$end))) {
        if (-not (Test-IPv4String $pair[1])) { throw "Invalid $($pair[0]) IPv4 address: $($pair[1])" }
        if (-not (Test-IPInCidr $pair[1] $cidr)) { throw "$($pair[0]) $($pair[1]) is not inside CIDR $cidr" }
    }
    if ((Convert-IPv4ToUInt32 $start) -gt (Convert-IPv4ToUInt32 $end)) { throw 'TEP Range Start must be less than or equal to Range End.' }
    $vlan = 0
    if (-not [int]::TryParse($vlanText, [ref]$vlan) -or $vlan -lt 0 -or $vlan -gt 4094) { throw 'Transport VLAN must be an integer from 0 to 4094.' }
}

# =============================================================================
# API helpers
# =============================================================================
function Get-HttpErrorDetail {
    param([object]$Ex)
    $code = ''
    $snippet = ''
    try {
        if ($Ex.Response -is [System.Net.Http.HttpResponseMessage]) {
            $code = [int]$Ex.Response.StatusCode
            $snippet = $Ex.Response.Content.ReadAsStringAsync().Result
        } elseif ($Ex.Response) {
            $code = $Ex.Response.StatusCode.value__
        }
    } catch {}
    if (-not $snippet) { $snippet = $Ex.Message }
    if ($snippet -and $snippet.Length -gt 2000) { $snippet = $snippet.Substring(0,2000) }
    return [pscustomobject]@{ Code=$code; Snippet=$snippet }
}

function Invoke-SddcApi {
    param([psobject]$Session,[ValidateSet('GET','POST','PATCH','PUT','DELETE')][string]$Method,[string]$Path,[object]$Body=$null,[switch]$ReturnHeaders)
    $uri = "$($Session.Base)$Path"
    $headers = @{ Authorization = "Bearer $($Session.AccessToken)"; Accept='application/json' }
    try {
        if ($ReturnHeaders) {
            $rh = $null
            if ($null -ne $Body) {
                $json = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 100 }
                $result = Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType 'application/json' -Body $json -SkipCertificateCheck -ResponseHeadersVariable rh
            } else {
                $result = Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType 'application/json' -SkipCertificateCheck -ResponseHeadersVariable rh
            }
            return [pscustomobject]@{ Body=$result; Headers=$rh }
        }
        if ($Method -in @('POST','PATCH','PUT')) {
            if ($null -ne $Body) {
                $json = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 100 }
                return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType 'application/json' -Body $json -SkipCertificateCheck
            }
            return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType 'application/json' -SkipCertificateCheck
        }
        return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -SkipCertificateCheck
    } catch {
        $d = Get-HttpErrorDetail $_.Exception
        throw "SDDC Manager API call failed: $Method $Path -> HTTP $($d.Code) $($d.Snippet)"
    }
}

function New-SddcToken {
    param([string]$SddcHost,[string]$Username,[string]$Password)
    $base = "https://$SddcHost"
    $body = @{ username=$Username; password=$Password } | ConvertTo-Json
    $tok = Invoke-RestMethod -Method POST -Uri "$base/v1/tokens" -ContentType 'application/json' -Body $body -SkipCertificateCheck
    if (-not $tok.accessToken) { throw 'Token response missing accessToken.' }
    return [pscustomobject]@{ Base=$base; AccessToken=$tok.accessToken; Host=$SddcHost; User=$Username }
}

function Invoke-NsxApi {
    param([psobject]$Session,[string]$Path)
    $pair = '{0}:{1}' -f $Session.User, $Session.Password
    $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
    $headers = @{ Authorization = "Basic $b64"; Accept='application/json' }
    try { return Invoke-RestMethod -Method GET -Uri "$($Session.Base)$Path" -Headers $headers -SkipCertificateCheck }
    catch { $d = Get-HttpErrorDetail $_.Exception; throw "NSX API call failed: GET $Path -> HTTP $($d.Code) $($d.Snippet)" }
}

function New-NsxSession {
    param([string]$NsxHost,[string]$Username,[string]$Password)
    $s = [pscustomobject]@{ Base="https://$NsxHost"; Host=$NsxHost; User=$Username; Password=$Password }
    $null = Invoke-NsxApi -Session $s -Path '/policy/api/v1/infra'
    return $s
}

function Get-Elements {
    param([object]$Response,[string[]]$PropertyNames)
    if ($null -eq $Response) { return @() }
    if ($Response.elements) { return @($Response.elements) }
    foreach ($p in $PropertyNames) { if ($Response.PSObject.Properties.Name -contains $p -and $Response.$p) { return @($Response.$p) } }
    return @($Response)
}

function Get-Clusters { param([psobject]$Session) Get-Elements -Response (Invoke-SddcApi -Session $Session -Method GET -Path '/v1/clusters') -PropertyNames @('clusters') }
function Get-Hosts { param([psobject]$Session) Get-Elements -Response (Invoke-SddcApi -Session $Session -Method GET -Path '/v1/hosts') -PropertyNames @('hosts') }

function Get-HostIdByFqdn {
    param([psobject]$Session,[string]$Fqdn)
    $all = @(Get-Hosts -Session $Session)
    $m = $all | Where-Object { ([string]$_.fqdn) -ieq $Fqdn -or ([string]$_.hostname) -ieq $Fqdn -or ([string]$_.FQDN) -ieq $Fqdn } | Select-Object -First 1
    if ($m) {
        $id = @($m.id,$m.hostId) | Where-Object { $_ } | Select-Object -First 1
        if ($id) { return ([string]$id) }
    }
    return $null
}

# =============================================================================
# Flatten / NSX inventory
# =============================================================================
function ConvertTo-FlatPairs {
    param([object]$Obj,[string]$Prefix='')
    $pairs = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Obj) { return $pairs }
    if ($Obj -is [System.Collections.IDictionary]) {
        foreach ($k in $Obj.Keys) {
            $p = if ($Prefix) { "$Prefix.$k" } else { [string]$k }
            foreach ($c in (ConvertTo-FlatPairs -Obj $Obj[$k] -Prefix $p)) { $pairs.Add($c) }
        }
        return $pairs
    }
    if ($Obj -is [pscustomobject]) {
        foreach ($pr in $Obj.PSObject.Properties) {
            $p = if ($Prefix) { "$Prefix.$($pr.Name)" } else { [string]$pr.Name }
            foreach ($c in (ConvertTo-FlatPairs -Obj $pr.Value -Prefix $p)) { $pairs.Add($c) }
        }
        return $pairs
    }
    if ($Obj -is [System.Collections.IEnumerable] -and $Obj -isnot [string]) {
        $i = 0
        foreach ($item in $Obj) {
            $p = "{0}[{1}]" -f $Prefix,$i
            foreach ($c in (ConvertTo-FlatPairs -Obj $item -Prefix $p)) { $pairs.Add($c) }
            $i++
        }
        return $pairs
    }
    $pairs.Add([pscustomobject]@{ Path=$Prefix; Value=$Obj })
    return $pairs
}

function Get-NsxIpPoolRanges {
    param([psobject]$Session)
    $responses = @()
    foreach ($path in @('/policy/api/v1/infra/ip-pools','/api/v1/pools/ip-pools')) {
        try { $responses += Invoke-NsxApi -Session $Session -Path $path }
        catch { Write-Log "NSX pool endpoint unavailable: $path" 'WARN' }
    }
    $ranges = @()
    $names = @()
    foreach ($r in $responses) {
        $flat = @(ConvertTo-FlatPairs -Obj $r)
        $names += @($flat | Where-Object { $_.Path -match '(display_name|name|id)$' } | ForEach-Object { [string]$_.Value } | Where-Object { $_ })
        $starts = @($flat | Where-Object { $_.Path -match '(start|start_ip|startAddress)$' } | ForEach-Object { [string]$_.Value } | Where-Object { Test-IPv4String $_ })
        $ends = @($flat | Where-Object { $_.Path -match '(end|end_ip|endAddress)$' } | ForEach-Object { [string]$_.Value } | Where-Object { Test-IPv4String $_ })
        for ($i=0; $i -lt [Math]::Min($starts.Count,$ends.Count); $i++) {
            $ranges += [pscustomobject]@{ Start=$starts[$i]; End=$ends[$i]; StartInt=(Convert-IPv4ToUInt32 $starts[$i]); EndInt=(Convert-IPv4ToUInt32 $ends[$i]) }
        }
    }
    return [pscustomobject]@{ Ranges=$ranges; Names=($names | Select-Object -Unique) }
}

function Test-NsxAvailability {
    param([switch]$ThrowOnFail)
    Assert-Az2PoolFields
    if (-not $script:NsxSession) {
        $msg = 'NSX Manager is not connected. NSX IP overlap check skipped.'
        Write-Log $msg 'WARN'
        if ($ThrowOnFail) { throw $msg }
        return [pscustomobject]@{ Ok=$true; Message=$msg }
    }
    $pool = ([string]$script:txtTepPoolName.Text).Trim()
    $cidr = ([string]$script:txtTepCidr.Text).Trim()
    $gw = ([string]$script:txtTepGateway.Text).Trim()
    $start = ([string]$script:txtTepRangeStart.Text).Trim()
    $end = ([string]$script:txtTepRangeEnd.Text).Trim()
    $s = Convert-IPv4ToUInt32 $start
    $e = Convert-IPv4ToUInt32 $end
    $data = Get-NsxIpPoolRanges -Session $script:NsxSession
    $over = @($data.Ranges | Where-Object { Test-RangeOverlap $s $e $_.StartInt $_.EndInt })
    if ($over.Count -gt 0) {
        $needed = [uint32]($e - $s + 1)
        $suggest = Find-AvailableRange -Cidr $cidr -Needed $needed -Used $data.Ranges -Gateway $gw
        $msg = "Requested TEP range $start-$end overlaps existing NSX range(s): $((@($over | ForEach-Object { "$($_.Start)-$($_.End)" })) -join ', ')."
        if ($suggest) { $msg += " Suggested available range: $($suggest.Start)-$($suggest.End)." }
        Write-Log $msg 'ERROR'
        if ($ThrowOnFail) { throw $msg }
        return [pscustomobject]@{ Ok=$false; Message=$msg }
    }
    $msg = "NSX IP availability check passed for $start-$end."
    if (@($data.Names | Where-Object { $_ -eq $pool }).Count -gt 0) { $msg += " Note: an NSX pool named '$pool' already exists." }
    Write-Log $msg 'INFO'
    return [pscustomobject]@{ Ok=$true; Message=$msg }
}

function Get-NsxUplinkProfiles {
    param([psobject]$Session)
    $profiles = @()
    foreach ($path in @('/policy/api/v1/infra/host-switch-profiles','/api/v1/host-switch-profiles')) {
        try {
            $resp = Invoke-NsxApi -Session $Session -Path $path
            if ($resp.results) { $profiles += @($resp.results) }
            elseif ($resp.elements) { $profiles += @($resp.elements) }
            elseif ($resp.host_switch_profiles) { $profiles += @($resp.host_switch_profiles) }
            elseif ($resp) { $profiles += @($resp) }
        } catch { Write-Log "NSX uplink profile endpoint unavailable: $path" 'WARN' }
    }
    $uplinkProfiles = @()
    foreach ($p in $profiles) {
        $text = ($p | ConvertTo-Json -Depth 30 -Compress)
        if ($text -match 'UplinkHostSwitchProfile|uplink') { $uplinkProfiles += $p }
    }
    return @($uplinkProfiles)
}

function Get-NsxUplinkProfileInventory {
    param([psobject]$Session)
    $profiles = @(Get-NsxUplinkProfiles -Session $Session)
    $names = @()
    $uplinks = @()
    $vlans = @()
    $policies = @()
    foreach ($p in $profiles) {
        foreach ($nameProp in @('display_name','displayName','name','id')) {
            try {
                $raw = $p.$nameProp
                $v = ([string]$raw).Trim()
                if ($v) { $names += $v }
            } catch {}
        }
        $flat = @(ConvertTo-FlatPairs -Obj $p)
        foreach ($f in $flat) {
            $path = [string]$f.Path
            $value = ([string]$f.Value).Trim()
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            if ($path -match '(active|standby|standBy).*uplink' -or $path -match 'uplink_name|uplinkName|uplinks\[\d+\]$') {
                if ($value -notmatch '^\d+$' -and $value.Length -lt 80) { $uplinks += $value }
            }
            if ($path -match 'transport.*vlan|transport_vlan|transportVlan|vlan$') {
                if ($value -match '^\d+$') { $vlans += $value }
            }
            if ($path -match 'policy|teaming') {
                if ($value -match 'LOADBALANCE|FAILOVER|ORDER|SRC') { $policies += $value }
            }
        }
    }
    return [pscustomobject]@{
        Profiles = $profiles
        Names = @($names | Select-Object -Unique)
        Uplinks = @($uplinks | Select-Object -Unique)
        Vlans = @($vlans | Select-Object -Unique)
        Policies = @($policies | Select-Object -Unique)
    }
}

function Test-NsxUplinkProfileValidation {
    param([string]$UplinkProfileName,[string]$ActiveUplinksCsv,[string]$StandbyUplinksCsv,[string]$TransportVlan,[string]$TeamingPolicy,[switch]$ThrowOnFail)
    if (-not $script:NsxSession) {
        Write-Log 'NSX Manager is not connected. NSX uplink profile validation skipped.' 'WARN'
        return [pscustomobject]@{ Ok=$true; Message='NSX uplink profile validation skipped.' }
    }
    $inv = Get-NsxUplinkProfileInventory -Session $script:NsxSession
    if (-not $inv.Profiles -or $inv.Profiles.Count -eq 0) {
        $msg = 'NSX uplink profile validation could not find uplink profiles from NSX APIs. Continuing with script-generated values.'
        Write-Log $msg 'WARN'
        return [pscustomobject]@{ Ok=$true; Message=$msg }
    }
    $messages = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    if ($UplinkProfileName -and ($inv.Names -contains $UplinkProfileName)) {
        $messages.Add("Generated uplink profile name '$UplinkProfileName' already exists in NSX. SDDC validation may reuse or reject it depending on workflow behavior.") | Out-Null
    }
    $requested = @((Parse-CommaList $ActiveUplinksCsv) + (Parse-CommaList $StandbyUplinksCsv)) | Where-Object { $_ } | Select-Object -Unique
    if ($inv.Uplinks -and $inv.Uplinks.Count -gt 0) {
        foreach ($u in $requested) {
            if ($inv.Uplinks -notcontains $u) { $errors.Add("Requested NSX uplink '$u' was not found in NSX-known uplink names: $($inv.Uplinks -join ', ')") | Out-Null }
        }
    } else {
        $messages.Add('NSX API returned uplink profiles, but no active/standby uplink names could be extracted. Uplink-name validation skipped.') | Out-Null
    }
    if ($TransportVlan -and $inv.Vlans -and $inv.Vlans.Count -gt 0 -and ($inv.Vlans -notcontains ([string]$TransportVlan))) {
        $messages.Add("Transport VLAN '$TransportVlan' was not found in existing NSX uplink profile VLAN values: $($inv.Vlans -join ', ')") | Out-Null
    }
    if ($TeamingPolicy -and $inv.Policies -and $inv.Policies.Count -gt 0 -and ($inv.Policies -notcontains $TeamingPolicy)) {
        $messages.Add("Teaming policy '$TeamingPolicy' was not found in extracted NSX profile policies: $($inv.Policies -join ', ')") | Out-Null
    }
    foreach ($m in $messages) { Write-Log $m 'WARN' }
    if ($errors.Count -gt 0) {
        $msg = 'NSX uplink profile validation failed. ' + ($errors -join ' ')
        Write-Log $msg 'ERROR'
        if ($ThrowOnFail) { throw $msg }
        return [pscustomobject]@{ Ok=$false; Message=$msg }
    }
    $okMsg = 'NSX uplink profile validation passed.'
    if ($messages.Count -gt 0) { $okMsg += ' Warnings were logged for operator review.' }
    Write-Log $okMsg 'INFO'
    return [pscustomobject]@{ Ok=$true; Message=$okMsg }
}

# =============================================================================
# Detection / spec helpers
# =============================================================================
function Get-SelectedClusterId {
    if (-not $script:cmbCluster.SelectedItem) { return $null }
    $label = $script:cmbCluster.SelectedItem.ToString()
    if ($script:ClusterMap.ContainsKey($label)) { return $script:ClusterMap[$label] }
    return $null
}

function Get-SelectedClusterName {
    if (-not $script:cmbCluster.SelectedItem) { return '' }
    $label = $script:cmbCluster.SelectedItem.ToString()
    if ($label -match '^(.+?)\s*\(') { return $Matches[1].Trim() }
    return $label
}

function Get-VCenterClusterNetworkFallback {
    $clusterName = Get-SelectedClusterName
    $cluster = Get-Cluster -Name $clusterName -ErrorAction Stop
    $vmhosts = @(Get-VMHost -Location $cluster -ErrorAction Stop)
    $rows = @()
    foreach ($h in $vmhosts) {
        $hv = Get-View -Id $h.Id -ErrorAction Stop
        foreach ($ps in @($hv.Config.Network.ProxySwitch)) {
            $vds = ([string]$ps.DvsName).Trim()
            if ($vds) { $rows += [pscustomobject]@{ vmhost=$h.Name; vdsName=$vds; pnicDevices=@($ps.Pnic) } }
        }
    }
    return [pscustomobject]@{ source='vCenter PowerCLI host proxy switches'; vdsName=@($rows | Select-Object -ExpandProperty vdsName -Unique); hostProxySwitches=@($rows) }
}

function New-DetectedNetworkObject {
    return [pscustomobject]@{ ClusterId=''; ClusterName=''; VdsCount=0; VdsNames=@(); PrimaryVdsName=''; VdsToNsxUplinkMappings=@(); ActiveUplinks=@(); StandbyUplinks=@(); TransportVlan=$null; TeamingPolicy='LOADBALANCE_SRCID' }
}

function Find-FlatValue {
    param([object[]]$FlatPairs,[string[]]$PathRegex,[string]$Mode='First')
    $hits = foreach ($re in $PathRegex) { $FlatPairs | Where-Object { $_.Path -match $re } | Select-Object -ExpandProperty Value }
    $hits = @($hits) | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace(([string]$_)) }
    if ($Mode -eq 'All') { return $hits }
    return ($hits | Select-Object -First 1)
}

function Update-DetectedNetworkSummary {
    param($DetectedNetwork)
    try {
        $script:lblDetectedVdsCount.Text = ([string]$DetectedNetwork.VdsCount)
        $script:lblDetectedVdsNames.Text = ($DetectedNetwork.VdsNames -join ', ')
        $map = (($DetectedNetwork.VdsToNsxUplinkMappings | ForEach-Object { "$($_.vdsUplinkName)->$($_.nsxUplinkName)" }) -join '; ')
        $script:lblDetectedUplinks.Text = $map
        $script:txtVdsToNsxUplinkMap.Text = $map
        $script:txtActiveUplinks.Text = ($DetectedNetwork.ActiveUplinks -join ',')
        $script:txtStandbyUplinks.Text = ($DetectedNetwork.StandbyUplinks -join ',')
    } catch {}
}

function Ensure-DefaultNetworkMappings {
    param([switch]$Force)
    if (-not $script:SddcSession) { throw 'Connect to SDDC Manager first.' }
    $cid = Get-SelectedClusterId
    if (-not $cid) { throw 'Select a cluster.' }
    if ($script:DetectedNetwork -and -not $Force -and $script:DetectedNetwork.ClusterId -eq $cid) { return $script:DetectedNetwork }
    Write-Log "Detecting network model for cluster $cid..."
    Write-Log 'VCF network query endpoint unavailable; using vCenter/cluster fallback detection.' 'INFO'

    $results = [ordered]@{}
    try { $results[("/v1/clusters/{0}" -f $cid)] = Invoke-SddcApi -Session $script:SddcSession -Method GET -Path ("/v1/clusters/{0}" -f $cid) } catch {}
    try { if ($script:VCenterVerified) { $results['vCenterPowerCLI'] = Get-VCenterClusterNetworkFallback } } catch { Write-Log "vCenter fallback detection failed: $($_.Exception.Message)" 'WARN' }
    if ($results.Count -eq 0) { throw 'No fallback network data available.' }
    $qr = [pscustomobject]$results
    if ($script:RunDir) {
        $raw = Join-Path $script:RunDir ("DetectedNetworkRaw_{0}_{1}.json" -f $cid,(Get-Date -Format 'yyyyMMdd-HHmmss'))
        $qr | ConvertTo-Json -Depth 90 | Set-Content -Path $raw -Encoding UTF8
        Write-Log "Saved raw detected network data: $raw"
    }
    $flat = @(ConvertTo-FlatPairs -Obj $qr)
    $vds = @(Find-FlatValue -FlatPairs $flat -PathRegex @('\.vdsName$','\.dvsName$','\.distributedSwitchName$') -Mode All) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -and $_ -notmatch '^[vV][dD][sS]0?\d$' } | Select-Object -Unique
    $det = New-DetectedNetworkObject
    $det.ClusterId = $cid
    $det.ClusterName = Get-SelectedClusterName
    $det.VdsNames = @($vds)
    $det.VdsCount = $det.VdsNames.Count
    if ($det.VdsCount -lt 1) { throw 'Unable to detect full vDS name.' }
    $det.PrimaryVdsName = $det.VdsNames[0]
    $det.VdsToNsxUplinkMappings = @([pscustomobject]@{vdsUplinkName='uplink1';nsxUplinkName='uplink1'},[pscustomobject]@{vdsUplinkName='uplink2';nsxUplinkName='uplink2'})
    $det.ActiveUplinks = @('uplink1','uplink2')
    $script:DetectedNetwork = $det
    Update-DetectedNetworkSummary -DetectedNetwork $det
    Write-Log "Detected vDS: $($det.VdsNames -join ', ')"
    return $det
}

function Assert-DetectedNetworkIsUsable { param($DetectedNetwork) if (-not $DetectedNetwork.VdsNames -or $DetectedNetwork.VdsNames.Count -lt 1) { throw 'Missing detected vDS names.' } }
function Parse-CommaList { param([string]$Text) if ([string]::IsNullOrWhiteSpace($Text)) { return @() }; return @($Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
function Parse-VmnicMapping { param([string]$Text) $out=@(); foreach($c in ($Text -split ';')){ $c=$c.Trim(); if(-not $c){continue}; if($c -match '^(vmnic\d+)\s*->\s*(.+?)\s*/\s*(uplink\S+)\s*$'){ $out += [pscustomobject]@{ id=$Matches[1]; vdsName=$Matches[2].Trim(); uplink=$Matches[3].Trim() } } else { throw "Invalid vmnic mapping: $c" } }; return $out }
function Parse-UplinkMapping { param([string]$Text) $out=@(); foreach($c in ($Text -split ';')){ $c=$c.Trim(); if(-not $c){continue}; if($c -match '^(.+?)\s*->\s*(.+?)\s*$'){ $out += [pscustomobject]@{ vdsUplinkName=$Matches[1].Trim(); nsxUplinkName=$Matches[2].Trim() } } }; return $out }
function Convert-DetectedUplinkMapToText { param($Mappings) return (($Mappings | ForEach-Object { "$($_.vdsUplinkName)->$($_.nsxUplinkName)" }) -join '; ') }
function Convert-DetectedVmnicMapToText { param($DetectedNetwork) $v1=$DetectedNetwork.VdsNames[0]; if($DetectedNetwork.VdsCount -le 1){return "vmnic0->$v1/uplink1; vmnic1->$v1/uplink2"}; $v2=$DetectedNetwork.VdsNames[1]; return "vmnic0->$v1/uplink1; vmnic1->$v1/uplink2; vmnic2->$v2/uplink1; vmnic3->$v2/uplink2" }
function Get-GeneratedAz2NetworkProfileName { $az2=([string]$script:txtAz2.Text).Trim(); if(-not $az2){$az2='AZ2'}; return "$az2-network-profile-01" }
function Get-GeneratedUplinkProfileName { return "$(Get-GeneratedAz2NetworkProfileName)-uplink-profile-01" }
function Get-ComboContent { param($Combo) if ($Combo -and $Combo.SelectedItem) { return ([string]$Combo.SelectedItem.Content).Trim() }; return '' }

function New-ClusterStretchSpec {
    param($Session,$ClusterId,$Az2Name,$Az2HostsFqdn,$Az2NetworkProfileName,$VmnicMappingText,$NsxHostSwitchVdsName,$TepPoolName,$TepCidr,$TepGateway,$TepRangeStart,$TepRangeEnd,$UplinkProfileName,[int]$TransportVlan,$TeamingPolicy,$ActiveUplinksCsv,$StandbyUplinksCsv,$VdsToNsxUplinkMapText,[bool]$DeployWithoutLicenseKeys,[bool]$IsEdgeClusterConfiguredForMultiAZ,$WitnessFqdn,$WitnessVsanIp,$WitnessVsanCidr,[bool]$WitnessTrafficSharedWithVsanTraffic)
    $vmnics = @(Parse-VmnicMapping $VmnicMappingText)
    $uplinkMap = @(Parse-UplinkMapping $VdsToNsxUplinkMapText)
    $hostSpecs = @()
    foreach ($fqdn in $Az2HostsFqdn) {
        $id = Get-HostIdByFqdn -Session $Session -Fqdn $fqdn
        if (-not $id) { throw "Could not resolve host ID for $fqdn" }
        $hostSpecs += [pscustomobject]@{ id=$id; hostname=$fqdn; azName=$Az2Name; hostNetworkSpec=[pscustomobject]@{ networkProfileName=$Az2NetworkProfileName; vmNics=@($vmnics) } }
    }
    return [pscustomobject]@{ clusterStretchSpec=[pscustomobject]@{
        deployWithoutLicenseKeys=$DeployWithoutLicenseKeys
        hostSpecs=@($hostSpecs)
        networkSpec=[pscustomobject]@{
            networkProfiles=@([pscustomobject]@{ isDefault=$true; name=$Az2NetworkProfileName; nsxtHostSwitchConfigs=@([pscustomobject]@{ ipAddressPoolName=$TepPoolName; uplinkProfileName=$UplinkProfileName; vdsName=$NsxHostSwitchVdsName; vdsUplinkToNsxUplink=@($uplinkMap) }) })
            nsxClusterSpec=[pscustomobject]@{
                ipAddressPoolsSpec=@([pscustomobject]@{ name=$TepPoolName; subnets=@([pscustomobject]@{ cidr=$TepCidr; gateway=$TepGateway; ipAddressPoolRanges=@([pscustomobject]@{ start=$TepRangeStart; end=$TepRangeEnd }) }) })
                uplinkProfiles=@([pscustomobject]@{ name=$UplinkProfileName; transportVlan=$TransportVlan; teamings=@([pscustomobject]@{ name='DEFAULT'; policy=$TeamingPolicy; standByUplinks=@(Parse-CommaList $StandbyUplinksCsv); activeUplinks=@(Parse-CommaList $ActiveUplinksCsv) }) })
            }
        }
        isEdgeClusterConfiguredForMultiAZ=$IsEdgeClusterConfiguredForMultiAZ
        witnessSpec=[pscustomobject]@{ fqdn=$WitnessFqdn; vsanCidr=$WitnessVsanCidr; vsanIp=$WitnessVsanIp; witnessTrafficSharedWithVsanTraffic=$WitnessTrafficSharedWithVsanTraffic }
    }}
}

function Build-SpecFromUi {
    if (-not $script:SddcSession) { throw 'Connect to SDDC Manager first.' }
    if ($script:chkRequireVCenterVerify.IsChecked -and -not $script:VCenterVerified) { throw 'vCenter verification is required.' }
    $cid = Get-SelectedClusterId
    $det = Ensure-DefaultNetworkMappings -Force
    Assert-DetectedNetworkIsUsable $det
    Assert-Az2PoolFields
    $nsxIp = Test-NsxAvailability -ThrowOnFail:$false
    if ($script:NsxSession -and -not $nsxIp.Ok) { throw $nsxIp.Message }
    $az2 = ([string]$script:txtAz2.Text).Trim()
    $hosts = @(([string]$script:txtAz2Hosts.Text) -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $az2 -or $hosts.Count -lt 1) { throw 'AZ2 name and hosts are required.' }
    $up = ([string]$script:txtUplinkProfileName.Text).Trim()
    if (-not $up) { $up = Get-GeneratedUplinkProfileName; $script:txtUplinkProfileName.Text = $up }
    $map = ([string]$script:txtVdsToNsxUplinkMap.Text).Trim()
    if (-not $map) { $map = Convert-DetectedUplinkMapToText $det.VdsToNsxUplinkMappings; $script:txtVdsToNsxUplinkMap.Text = $map }
    $active = ([string]$script:txtActiveUplinks.Text).Trim()
    if (-not $active) { $active = ($det.VdsToNsxUplinkMappings | Select-Object -ExpandProperty nsxUplinkName -Unique) -join ','; $script:txtActiveUplinks.Text = $active }
    $standby = ([string]$script:txtStandbyUplinks.Text).Trim()
    $transportVlan = ([string]$script:txtTransportVlan.Text).Trim()
    $teamPolicy = Get-ComboContent $script:cmbTeamingPolicy
    if ($script:NsxSession) { $null = Test-NsxUplinkProfileValidation -UplinkProfileName $up -ActiveUplinksCsv $active -StandbyUplinksCsv $standby -TransportVlan $transportVlan -TeamingPolicy $teamPolicy -ThrowOnFail }
    return (New-ClusterStretchSpec -Session $script:SddcSession -ClusterId $cid -Az2Name $az2 -Az2HostsFqdn $hosts -Az2NetworkProfileName (Get-GeneratedAz2NetworkProfileName) -VmnicMappingText (Convert-DetectedVmnicMapToText $det) -NsxHostSwitchVdsName $det.PrimaryVdsName -TepPoolName (([string]$script:txtTepPoolName.Text).Trim()) -TepCidr (([string]$script:txtTepCidr.Text).Trim()) -TepGateway (([string]$script:txtTepGateway.Text).Trim()) -TepRangeStart (([string]$script:txtTepRangeStart.Text).Trim()) -TepRangeEnd (([string]$script:txtTepRangeEnd.Text).Trim()) -UplinkProfileName $up -TransportVlan ([int]$transportVlan) -TeamingPolicy $teamPolicy -ActiveUplinksCsv $active -StandbyUplinksCsv $standby -VdsToNsxUplinkMapText $map -DeployWithoutLicenseKeys ([bool]$script:chkDeployNoLic.IsChecked) -IsEdgeClusterConfiguredForMultiAZ ([bool]$script:chkEdgeMultiAZ.IsChecked) -WitnessFqdn (([string]$script:txtWitnessFqdn.Text).Trim()) -WitnessVsanIp (([string]$script:txtWitnessVsanIp.Text).Trim()) -WitnessVsanCidr (([string]$script:txtWitnessVsanCidr.Text).Trim()) -WitnessTrafficSharedWithVsanTraffic ([bool]$script:chkWitnessShared.IsChecked))
}

# =============================================================================
# UI
# =============================================================================
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase -ErrorAction Stop | Out-Null
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue | Out-Null

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="VCF 9 Stretch Cluster Automation v$Global:VCFStretchVersion" Height="980" Width="1600" Background="#1F2733" Foreground="#E6EDF3" FontFamily="Segoe UI" FontSize="12" WindowStartupLocation="CenterScreen">
<Window.Resources>
<Style TargetType="GroupBox"><Setter Property="Foreground" Value="#E6EDF3"/><Setter Property="Margin" Value="8"/><Setter Property="Padding" Value="6"/><Setter Property="BorderBrush" Value="#465568"/></Style>
<Style TargetType="TextBlock"><Setter Property="Foreground" Value="#E6EDF3"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
<Style TargetType="TextBox"><Setter Property="Background" Value="#131A23"/><Setter Property="Foreground" Value="#E6EDF3"/><Setter Property="BorderBrush" Value="#53657A"/><Setter Property="Margin" Value="4"/></Style>
<Style TargetType="PasswordBox"><Setter Property="Background" Value="#131A23"/><Setter Property="Foreground" Value="#E6EDF3"/><Setter Property="BorderBrush" Value="#53657A"/><Setter Property="Margin" Value="4"/></Style>
<Style TargetType="Button"><Setter Property="Margin" Value="4"/><Setter Property="Padding" Value="8,3"/></Style>
<Style TargetType="CheckBox"><Setter Property="Foreground" Value="#E6EDF3"/></Style>
</Window.Resources>
<Grid Margin="10">
<Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<TextBlock Grid.Row="0" Text="VCF 9 Stretch Cluster Automation" FontSize="24" FontWeight="Bold"/>
<GroupBox Grid.Row="1" Header="Prerequisites"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="200"/><ColumnDefinition Width="200"/><ColumnDefinition Width="240"/><ColumnDefinition Width="240"/><ColumnDefinition Width="240"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Name="lblPS" Grid.Column="0"/><TextBlock Name="lblWPF" Grid.Column="1"/><TextBlock Name="lblImpExcel" Grid.Column="2"/><TextBlock Name="lblPCLI" Grid.Column="3"/><TextBlock Name="lblVCFPCLI" Grid.Column="4"/><Button Name="btnRecheck" Grid.Column="5" Content="Recheck"/><Button Name="btnInstallPCLI" Grid.Column="6" Content="Install PowerCLI"/><Button Name="btnInstallVCFPCLI" Grid.Column="7" Content="Install VCF.PowerCLI"/></Grid></GroupBox>
<GroupBox Grid.Row="2" Header="Connections"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="310"/><ColumnDefinition Width="100"/><ColumnDefinition Width="270"/><ColumnDefinition Width="115"/><ColumnDefinition Width="170"/><ColumnDefinition Width="125"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><TextBlock Grid.Row="0" Grid.Column="0" Text="SDDC Manager:"/><TextBox Grid.Row="0" Grid.Column="1" Name="txtSddcHost"/><TextBlock Grid.Row="0" Grid.Column="2" Text="User:"/><TextBox Grid.Row="0" Grid.Column="3" Name="txtSddcUser" Text="administrator@vsphere.local"/><TextBlock Grid.Row="0" Grid.Column="4" Text="Password:"/><PasswordBox Grid.Row="0" Grid.Column="5" Name="pbSddcPass"/><Button Grid.Row="0" Grid.Column="6" Name="btnConnect" Content="Connect SDDC"/><TextBlock Grid.Row="0" Grid.Column="7" Name="lblConnStatus" Text="Not connected" Foreground="Gold"/><TextBlock Grid.Row="1" Grid.Column="0" Text="vCenter FQDN:"/><TextBox Grid.Row="1" Grid.Column="1" Name="txtVCenterFqdn"/><TextBlock Grid.Row="1" Grid.Column="2" Text="User:"/><TextBox Grid.Row="1" Grid.Column="3" Name="txtVCenterUser" Text="administrator@vsphere.local"/><TextBlock Grid.Row="1" Grid.Column="4" Text="Password:"/><PasswordBox Grid.Row="1" Grid.Column="5" Name="pbVCenterPass"/><Button Grid.Row="1" Grid.Column="6" Name="btnVerifyVCenter" Content="Verify vCenter"/><TextBlock Grid.Row="1" Grid.Column="7" Name="lblVCenterStatus" Text="Not verified" Foreground="Gold"/><TextBlock Grid.Row="2" Grid.Column="0" Text="NSX Manager:"/><TextBox Grid.Row="2" Grid.Column="1" Name="txtNsxHost"/><TextBlock Grid.Row="2" Grid.Column="2" Text="User:"/><TextBox Grid.Row="2" Grid.Column="3" Name="txtNsxUser" Text="admin"/><TextBlock Grid.Row="2" Grid.Column="4" Text="Password:"/><PasswordBox Grid.Row="2" Grid.Column="5" Name="pbNsxPass"/><Button Grid.Row="2" Grid.Column="6" Name="btnConnectNsx" Content="Connect NSX"/><TextBlock Grid.Row="2" Grid.Column="7" Name="lblNsxStatus" Text="NSX not connected" Foreground="Gold"/></Grid></GroupBox>
<GroupBox Grid.Row="3" Header="Cluster and Detection">
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="90"/>
      <ColumnDefinition Width="430"/>
      <ColumnDefinition Width="140"/>
      <ColumnDefinition Width="28"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Grid.Column="0" Text="Cluster:" VerticalAlignment="Center"/>
    <ComboBox Grid.Row="0" Grid.Column="1" Name="cmbCluster" MinWidth="390" Height="26" VerticalAlignment="Center"/>
    <Button Grid.Row="0" Grid.Column="2" Name="btnDetectNetwork" Content="Detect Network" Width="120" Height="26" VerticalAlignment="Center"/>

    <Border Grid.Row="0" Grid.Column="4" BorderBrush="#465568" BorderThickness="1" CornerRadius="3" Padding="8" Margin="8,0,0,0" Background="#18212B" VerticalAlignment="Center">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="95"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Grid.Column="0" Text="vDS Count:" FontWeight="SemiBold"/>
        <TextBlock Grid.Row="0" Grid.Column="1" Name="lblDetectedVdsCount" Text="0"/>
        <TextBlock Grid.Row="1" Grid.Column="0" Text="vDS Names:" FontWeight="SemiBold" Margin="0,3,0,0"/>
        <TextBlock Grid.Row="1" Grid.Column="1" Name="lblDetectedVdsNames" Text="Not detected" TextWrapping="Wrap" Margin="0,3,0,0"/>
        <TextBlock Grid.Row="2" Grid.Column="0" Text="Uplinks:" FontWeight="SemiBold" Margin="0,3,0,0"/>
        <TextBlock Grid.Row="2" Grid.Column="1" Name="lblDetectedUplinks" Text="Not detected" TextWrapping="Wrap" Margin="0,3,0,0"/>
      </Grid>
    </Border>
  </Grid>
</GroupBox><GroupBox Grid.Row="4" Header="AZ2 NSX TEP Pool"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="220"/><ColumnDefinition Width="80"/><ColumnDefinition Width="160"/><ColumnDefinition Width="90"/><ColumnDefinition Width="160"/><ColumnDefinition Width="110"/><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions><TextBlock Grid.Row="0" Grid.Column="0" Text="Pool Name:"/><TextBox Grid.Row="0" Grid.Column="1" Name="txtTepPoolName"/><TextBlock Grid.Row="0" Grid.Column="2" Text="CIDR:"/><TextBox Grid.Row="0" Grid.Column="3" Name="txtTepCidr"/><TextBlock Grid.Row="0" Grid.Column="4" Text="Gateway:"/><TextBox Grid.Row="0" Grid.Column="5" Name="txtTepGateway"/><TextBlock Grid.Row="0" Grid.Column="6" Text="TEP VLAN:"/><TextBox Grid.Row="0" Grid.Column="7" Name="txtTransportVlan"/><TextBlock Grid.Row="1" Grid.Column="0" Text="Range Start:"/><TextBox Grid.Row="1" Grid.Column="1" Name="txtTepRangeStart"/><TextBlock Grid.Row="1" Grid.Column="2" Text="Range End:"/><TextBox Grid.Row="1" Grid.Column="3" Name="txtTepRangeEnd"/><TextBlock Grid.Row="1" Grid.Column="4" Text="Summary:"/><TextBlock Grid.Row="1" Grid.Column="5" Grid.ColumnSpan="4" Name="lblAz2PoolSummary" Text="Pool/CIDR/Gateway/Range/VLAN required" Foreground="#9FB3C8"/></Grid></GroupBox>
<GroupBox Grid.Row="5" Header="Advanced Generated Names / Uplink Profile"><Expander Header="Show advanced values" Foreground="#E6EDF3"><Grid Margin="8"><Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="280"/><ColumnDefinition Width="110"/><ColumnDefinition Width="180"/><ColumnDefinition Width="110"/><ColumnDefinition Width="180"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/><RowDefinition/><RowDefinition/></Grid.RowDefinitions><TextBlock Grid.Row="0" Grid.Column="0" Text="AZ2 Profile:"/><TextBlock Grid.Row="0" Grid.Column="1" Name="lblGeneratedAz2Profile" Text="Auto" Foreground="#9FB3C8"/><TextBlock Grid.Row="0" Grid.Column="2" Text="Uplink Profile:"/><TextBox Grid.Row="0" Grid.Column="3" Grid.ColumnSpan="2" Name="txtUplinkProfileName"/><TextBlock Grid.Row="1" Grid.Column="0" Text="Teaming:"/><ComboBox Grid.Row="1" Grid.Column="1" Name="cmbTeamingPolicy"><ComboBoxItem Content="LOADBALANCE_SRCID" IsSelected="True"/><ComboBoxItem Content="FAILOVER_ORDER"/><ComboBoxItem Content="LOADBALANCE_SRC_MAC"/></ComboBox><TextBlock Grid.Row="1" Grid.Column="2" Text="Active:"/><TextBox Grid.Row="1" Grid.Column="3" Name="txtActiveUplinks"/><TextBlock Grid.Row="1" Grid.Column="4" Text="Standby:"/><TextBox Grid.Row="1" Grid.Column="5" Name="txtStandbyUplinks"/><TextBlock Grid.Row="2" Grid.Column="0" Text="vDS->NSX:"/><TextBox Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="5" Name="txtVdsToNsxUplinkMap"/></Grid></Expander></GroupBox>
<Grid Grid.Row="6"><Grid.ColumnDefinitions><ColumnDefinition Width="2*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><GroupBox Grid.Column="0" Header="Stretch Inputs"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="320"/><ColumnDefinition Width="120"/><ColumnDefinition Width="320"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition/><RowDefinition Height="150"/><RowDefinition/><RowDefinition/><RowDefinition/></Grid.RowDefinitions><TextBlock Grid.Row="0" Grid.Column="0" Text="AZ1 Name:"/><TextBox Grid.Row="0" Grid.Column="1" Name="txtAz1"/><TextBlock Grid.Row="0" Grid.Column="2" Text="AZ2 Name:"/><TextBox Grid.Row="0" Grid.Column="3" Name="txtAz2"/><TextBlock Grid.Row="1" Grid.Column="0" Text="AZ2 Hosts:" VerticalAlignment="Top"/><TextBox Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="3" Name="txtAz2Hosts" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/><Button Grid.Row="1" Grid.Column="4" Name="btnPickHosts" Content="Pick AZ2 Hosts" VerticalAlignment="Top"/><TextBlock Grid.Row="2" Grid.Column="0" Text="Witness FQDN:"/><TextBox Grid.Row="2" Grid.Column="1" Name="txtWitnessFqdn"/><TextBlock Grid.Row="2" Grid.Column="2" Text="Witness vSAN IP:"/><TextBox Grid.Row="2" Grid.Column="3" Name="txtWitnessVsanIp"/><TextBlock Grid.Row="3" Grid.Column="0" Text="Witness CIDR:"/><TextBox Grid.Row="3" Grid.Column="1" Name="txtWitnessVsanCidr"/><CheckBox Grid.Row="3" Grid.Column="2" Grid.ColumnSpan="2" Name="chkWitnessShared" Content="Witness traffic shared with vSAN traffic" IsChecked="True"/><CheckBox Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Name="chkDeployNoLic" Content="Deploy without license keys" IsChecked="True"/><CheckBox Grid.Row="4" Grid.Column="2" Grid.ColumnSpan="2" Name="chkEdgeMultiAZ" Content="Edge cluster configured for Multi-AZ"/><CheckBox Grid.Row="4" Grid.Column="4" Name="chkRequireVCenterVerify" Content="Require vCenter verification" IsChecked="True"/></Grid></GroupBox><GroupBox Grid.Column="1" Header="Log"><TextBox Name="txtLog" Background="#131A23" Foreground="#E6EDF3" FontFamily="Consolas" AcceptsReturn="True" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/></GroupBox></Grid>
<GroupBox Grid.Row="7" Header="Configuration / Output / Actions"><DockPanel LastChildFill="False"><StackPanel DockPanel.Dock="Left" Orientation="Horizontal"><TextBlock Text="Config:"/><TextBox Name="txtConfigPath" Width="340"/><Button Name="btnBrowseConfig" Content="Browse"/><Button Name="btnLoadConfig" Content="Load"/><Button Name="btnSaveConfig" Content="Save"/></StackPanel><StackPanel DockPanel.Dock="Right" Orientation="Horizontal"><TextBlock Text="Reports:"/><TextBox Name="txtReportsPath" Width="230"/><Button Name="btnBrowseReports" Content="Browse"/><Button Name="btnOpenOut" Content="Open Output"/><Button Name="btnGenerate" Content="Generate JSON"/><Button Name="btnValidate" Content="Validate"/><Button Name="btnExecute" Content="Execute"/><Button Name="btnClose" Content="Close"/></StackPanel></DockPanel></GroupBox>
</Grid></Window>
"@

$script:window = [Windows.Markup.XamlReader]::Parse($xaml)
foreach ($name in @('lblPS','lblWPF','lblImpExcel','lblPCLI','lblVCFPCLI','btnRecheck','btnInstallPCLI','btnInstallVCFPCLI','txtReportsPath','btnBrowseReports','btnOpenOut','txtSddcHost','txtSddcUser','pbSddcPass','btnConnect','lblConnStatus','txtVCenterFqdn','txtVCenterUser','pbVCenterPass','btnVerifyVCenter','lblVCenterStatus','txtNsxHost','txtNsxUser','pbNsxPass','btnConnectNsx','lblNsxStatus','cmbCluster','btnDetectNetwork','lblDetectedVdsCount','lblDetectedVdsNames','lblDetectedUplinks','txtTepPoolName','txtTepCidr','txtTepGateway','txtTepRangeStart','txtTepRangeEnd','txtTransportVlan','lblAz2PoolSummary','lblGeneratedAz2Profile','txtUplinkProfileName','cmbTeamingPolicy','txtActiveUplinks','txtStandbyUplinks','txtVdsToNsxUplinkMap','txtAz1','txtAz2','btnPickHosts','txtAz2Hosts','txtWitnessFqdn','txtWitnessVsanIp','txtWitnessVsanCidr','chkWitnessShared','chkDeployNoLic','chkEdgeMultiAZ','chkRequireVCenterVerify','txtConfigPath','btnBrowseConfig','btnLoadConfig','btnSaveConfig','txtLog','btnGenerate','btnValidate','btnExecute','btnClose')) { Set-Variable -Name $name -Scope Script -Value $script:window.FindName($name) }

function Select-ComboByContent { param($Combo,[string]$Content) foreach ($item in $Combo.Items) { if (([string]$item.Content) -eq $Content) { $Combo.SelectedItem = $item; return } } }
function Update-Az2PoolSummary { try { $pool=([string]$script:txtTepPoolName.Text).Trim(); $cidr=([string]$script:txtTepCidr.Text).Trim(); $gw=([string]$script:txtTepGateway.Text).Trim(); $start=([string]$script:txtTepRangeStart.Text).Trim(); $end=([string]$script:txtTepRangeEnd.Text).Trim(); $vlan=([string]$script:txtTransportVlan.Text).Trim(); if (@($pool,$cidr,$gw,$start,$end,$vlan) | Where-Object { [string]::IsNullOrWhiteSpace($_) }) { $script:lblAz2PoolSummary.Text='Pool/CIDR/Gateway/Range/VLAN required' } else { $script:lblAz2PoolSummary.Text="$pool | $cidr | GW $gw | $start-$end | VLAN $vlan" } } catch {} }
function Update-GeneratedNamesSummary { try { $script:lblGeneratedAz2Profile.Text = Get-GeneratedAz2NetworkProfileName; if ([string]::IsNullOrWhiteSpace(([string]$script:txtUplinkProfileName.Text))) { $script:txtUplinkProfileName.Text = Get-GeneratedUplinkProfileName } } catch {} }
function Refresh-TopologyUI { $hasSession=$script:SddcSession -ne $null; $hasCluster=$script:cmbCluster.SelectedItem -ne $null; $hasDetected=$script:DetectedNetwork -ne $null; $vcOk=(-not $script:chkRequireVCenterVerify.IsChecked) -or $script:VCenterVerified; $script:btnPickHosts.IsEnabled=$hasSession; $script:btnDetectNetwork.IsEnabled=$hasSession -and $hasCluster; $can=$hasSession -and $hasCluster -and $hasDetected -and $vcOk; $script:btnGenerate.IsEnabled=$can; $script:btnValidate.IsEnabled=$can; $script:btnExecute.IsEnabled=$can }
function Get-UiConfig { return [pscustomobject]@{ SddcHost=$script:txtSddcHost.Text; SddcUser=$script:txtSddcUser.Text; VCenterFqdn=$script:txtVCenterFqdn.Text; VCenterUser=$script:txtVCenterUser.Text; NsxHost=$script:txtNsxHost.Text; NsxUser=$script:txtNsxUser.Text; Az1=$script:txtAz1.Text; Az2=$script:txtAz2.Text; Az2Hosts=$script:txtAz2Hosts.Text; TepPoolName=$script:txtTepPoolName.Text; TepCidr=$script:txtTepCidr.Text; TepGateway=$script:txtTepGateway.Text; TepRangeStart=$script:txtTepRangeStart.Text; TepRangeEnd=$script:txtTepRangeEnd.Text; TransportVlan=$script:txtTransportVlan.Text; WitnessFqdn=$script:txtWitnessFqdn.Text; WitnessVsanIp=$script:txtWitnessVsanIp.Text; WitnessVsanCidr=$script:txtWitnessVsanCidr.Text } }
function Save-UiConfig { param($Path) (Get-UiConfig | ConvertTo-Json -Depth 8) | Set-Content -Path $Path -Encoding UTF8; Write-Log "Saved config: $Path" }
function Load-UiConfig { param($Path) $c=Get-Content -Path $Path -Raw | ConvertFrom-Json; foreach($p in $c.PSObject.Properties){$ctrl=Get-Variable -Name ('txt'+$p.Name) -Scope Script -ErrorAction SilentlyContinue; if($ctrl){$ctrl.Value.Text=$p.Value}}; Write-Log 'Config loaded into UI; passwords not loaded.' }

# Events
$script:window.Add_ContentRendered({ if(-not $script:RunDir){New-RunDir -Base $script:ReportsBase | Out-Null}; $script:txtReportsPath.Text=$script:ReportsBase; Select-ComboByContent $script:cmbTeamingPolicy 'LOADBALANCE_SRCID'; Prereq-Check | Out-Null; Write-Log "==== VCF Stretch UI started v$Global:VCFStretchVersion ===="; Write-Log "Run folder: $script:RunDir"; Refresh-TopologyUI })
$script:btnRecheck.Add_Click({ Prereq-Check | Out-Null })
$script:btnInstallPCLI.Add_Click({ try { Install-Module VMware.PowerCLI -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -AcceptLicense; Prereq-Check | Out-Null } catch { Write-Log "Install PowerCLI failed: $($_.Exception.Message)" ERROR } })
$script:btnInstallVCFPCLI.Add_Click({ try { Install-Module VCF.PowerCLI -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -AcceptLicense; Prereq-Check | Out-Null } catch { Write-Log "Install VCF.PowerCLI failed: $($_.Exception.Message)" ERROR } })
foreach($tb in @($script:txtTepPoolName,$script:txtTepCidr,$script:txtTepGateway,$script:txtTepRangeStart,$script:txtTepRangeEnd,$script:txtTransportVlan,$script:txtAz2)){ $tb.Add_TextChanged({ Update-Az2PoolSummary; Update-GeneratedNamesSummary }) }
$script:btnConnect.Add_Click({ try { $script:SddcSession=New-SddcToken -SddcHost (([string]$script:txtSddcHost.Text).Trim()) -Username (([string]$script:txtSddcUser.Text).Trim()) -Password ([string]$script:pbSddcPass.Password); $script:lblConnStatus.Text='Connected'; $script:lblConnStatus.Foreground=[Windows.Media.Brushes]::LightGreen; Write-Log "Connected to SDDC Manager: $($script:txtSddcHost.Text)"; $script:cmbCluster.Items.Clear(); $script:ClusterMap=@{}; foreach($c in @(Get-Clusters -Session $script:SddcSession)){ $cid=([string](@($c.id,$c.clusterId)|Where-Object{$_}|Select-Object -First 1)); $name=([string](@($c.name,$c.clusterName)|Where-Object{$_}|Select-Object -First 1)); if($cid){$label="$name ($cid)"; [void]$script:cmbCluster.Items.Add($label); $script:ClusterMap[$label]=$cid }}; if($script:cmbCluster.Items.Count -gt 0){$script:cmbCluster.SelectedIndex=0}; Refresh-TopologyUI } catch { Write-Log "Connect failed: $($_.Exception.Message)" ERROR; [System.Windows.MessageBox]::Show($_.Exception.Message,'Connect failed','OK','Error')|Out-Null } })
$script:btnConnectNsx.Add_Click({ try { $script:NsxSession=New-NsxSession -NsxHost (([string]$script:txtNsxHost.Text).Trim()) -Username (([string]$script:txtNsxUser.Text).Trim()) -Password ([string]$script:pbNsxPass.Password); $script:lblNsxStatus.Text='NSX connected'; $script:lblNsxStatus.Foreground=[Windows.Media.Brushes]::LightGreen; Write-Log "Connected to NSX Manager: $($script:txtNsxHost.Text)"; Write-Log 'NSX IP and uplink profile checks will run automatically before Generate, Validate, and Execute.' } catch { Write-Log "NSX connect failed: $($_.Exception.Message)" ERROR; [System.Windows.MessageBox]::Show($_.Exception.Message,'NSX connect failed','OK','Error')|Out-Null } })
$script:btnVerifyVCenter.Add_Click({ try { Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue | Out-Null; Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null; $cred=[pscredential]::new(([string]$script:txtVCenterUser.Text),(ConvertTo-SecureString ([string]$script:pbVCenterPass.Password) -AsPlainText -Force)); $script:VCenterServer=Connect-VIServer -Server (([string]$script:txtVCenterFqdn.Text).Trim()) -Credential $cred -WarningAction SilentlyContinue -ErrorAction Stop; $script:VCenterVerified=$true; $script:lblVCenterStatus.Text='Verified'; $script:lblVCenterStatus.Foreground=[Windows.Media.Brushes]::LightGreen; Write-Log "vCenter verified: $($script:txtVCenterFqdn.Text)"; Refresh-TopologyUI } catch { Write-Log "vCenter verification failed: $($_.Exception.Message)" ERROR; [System.Windows.MessageBox]::Show($_.Exception.Message,'vCenter failed','OK','Error')|Out-Null } })
$script:cmbCluster.Add_SelectionChanged({ $script:DetectedNetwork=$null; Refresh-TopologyUI })
$script:btnDetectNetwork.Add_Click({ try { $d=Ensure-DefaultNetworkMappings -Force; [System.Windows.MessageBox]::Show("Detected vDS: $($d.VdsNames -join ', ')",'Detect Network','OK','Information')|Out-Null; Refresh-TopologyUI } catch { Write-Log "Network detection failed: $($_.Exception.Message)" ERROR; [System.Windows.MessageBox]::Show($_.Exception.Message,'Detect failed','OK','Error')|Out-Null } })
$script:btnBrowseReports.Add_Click({ $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){$script:ReportsBase=$dlg.SelectedPath; $script:txtReportsPath.Text=$script:ReportsBase; New-RunDir -Base $script:ReportsBase | Out-Null} })
$script:btnOpenOut.Add_Click({ if($script:RunDir){Start-Process $script:RunDir} })
$script:btnBrowseConfig.Add_Click({ $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter='JSON config (*.json)|*.json'; if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){$script:txtConfigPath.Text=$dlg.FileName} })
$script:btnLoadConfig.Add_Click({ try { Load-UiConfig -Path $script:txtConfigPath.Text } catch { [System.Windows.MessageBox]::Show($_.Exception.Message) | Out-Null } })
$script:btnSaveConfig.Add_Click({ $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter='JSON config (*.json)|*.json'; $dlg.FileName='vcf-stretch-config.json'; if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){$script:txtConfigPath.Text=$dlg.FileName; Save-UiConfig -Path $dlg.FileName} })
$script:btnPickHosts.Add_Click({ try { $hosts=@(Get-Hosts -Session $script:SddcSession | ForEach-Object { @($_.fqdn,$_.hostname) | Where-Object { $_ } | Select-Object -First 1 } | Select-Object -Unique); $script:txtAz2Hosts.Text=($hosts -join "`r`n") } catch { Write-Log "Host picker failed: $($_.Exception.Message)" WARN } })
$script:btnGenerate.Add_Click({ try { $spec=Build-SpecFromUi; $ts=Get-Date -Format yyyyMMdd-HHmmss; $cid=Get-SelectedClusterId; $out=Join-Path $script:RunDir "clusterStretchSpec_$cid`_$ts.json"; $wrap=Join-Path $script:RunDir "clusterUpdateSpec_validationWrapper_$cid`_$ts.json"; $spec|ConvertTo-Json -Depth 90|Set-Content $out -Encoding UTF8; @{clusterUpdateSpec=$spec}|ConvertTo-Json -Depth 100|Set-Content $wrap -Encoding UTF8; Write-Log "Wrote JSON: $out"; Write-Log "Wrote validation wrapper: $wrap"; [System.Windows.MessageBox]::Show("Generated:`n$out`n`n$wrap")|Out-Null } catch { Write-Log "Generate failed: $($_.Exception.Message)" ERROR; [System.Windows.MessageBox]::Show($_.Exception.Message,'Generate failed','OK','Error')|Out-Null } })
$script:btnValidate.Add_Click({ try { $spec=Build-SpecFromUi; $cid=Get-SelectedClusterId; Write-Log "Validating stretch spec via POST /v1/clusters/$cid/validations..."; $res=Invoke-SddcApi -Session $script:SddcSession -Method POST -Path ("/v1/clusters/{0}/validations" -f $cid) -Body @{clusterUpdateSpec=$spec}; $out=Join-Path $script:RunDir ("ValidationResponse_{0}_{1}.json" -f $cid,(Get-Date -Format yyyyMMdd-HHmmss)); $res|ConvertTo-Json -Depth 90|Set-Content $out -Encoding UTF8; Write-Log "Validation response saved: $out"; [System.Windows.MessageBox]::Show("Validation submitted. Response saved:`n$out")|Out-Null } catch { Write-Log "Validate failed: $($_.Exception.Message)" ERROR; [System.Windows.MessageBox]::Show($_.Exception.Message,'Validate failed','OK','Error')|Out-Null } })
$script:btnExecute.Add_Click({ try { $spec=Build-SpecFromUi; $cid=Get-SelectedClusterId; $ans=[System.Windows.MessageBox]::Show("This will PATCH /v1/clusters/$cid and start stretch. Continue?",'Execute','YesNo','Warning'); if($ans -ne 'Yes'){return}; $res=Invoke-SddcApi -Session $script:SddcSession -Method PATCH -Path ("/v1/clusters/{0}" -f $cid) -Body $spec; $out=Join-Path $script:RunDir ("ExecuteResponse_{0}_{1}.json" -f $cid,(Get-Date -Format yyyyMMdd-HHmmss)); $res|ConvertTo-Json -Depth 90|Set-Content $out -Encoding UTF8; Write-Log "Execute response saved: $out"; [System.Windows.MessageBox]::Show("Execute submitted. Response saved:`n$out")|Out-Null } catch { Write-Log "Execute failed: $($_.Exception.Message)" ERROR; [System.Windows.MessageBox]::Show($_.Exception.Message,'Execute failed','OK','Error')|Out-Null } })
$script:btnClose.Add_Click({ $script:window.Close() })
$null = $script:window.ShowDialog()


