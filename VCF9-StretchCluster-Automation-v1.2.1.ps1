<#
.SYNOPSIS
VCF 9 Stretched Cluster Automation (ESA/OSA, 1 or 2 vDS)

.DESCRIPTION
WPF-based tool (derived from VCF9TrustedChain foundation UI patterns) that:
  * Collects required data for a VCF 9 stretched cluster spec using SDDC Manager API (source of truth)
  * Populates the provided Excel data-collection template ("Collected Data (replace example)" column)
  * Generates ready-to-use JSON payloads:
      - PATCH payload: { "clusterStretchSpec": { ... } }
      - Validation wrapper: { "clusterUpdateSpec": { "clusterStretchSpec": { ... } } }
  * Optionally validates (POST /v1/clusters/{id}/validations) and executes stretch (PATCH /v1/clusters/{id})

.NOTES
- Requires PowerShell 7+ for best experience.
- Uses ImportExcel module to edit the workbook without Excel installed.
- Uses SDDC Manager API endpoints (/v1/...) per VCF 9 guidance.

#>

[CmdletBinding()]
param(
  [switch]$NoRelaunch,
  [switch]$SignedOk,
  [switch]$NoAutoSign
)

$Global:VCFStretchVersion = '0.10.41'
$VerbosePreference='SilentlyContinue'
$InformationPreference='Continue'
$ProgressPreference='SilentlyContinue'

function Coalesce([object]$a,[object]$b){ if ($null -ne $a -and ($a -isnot [string] -or $a -ne '')) { $a } else { $b } }

# --- Self-sign + STA relaunch (kept consistent with foundation script patterns) ---
function Ensure-SelfSigned {
  param([string]$TargetPath)
  try { $sig = Get-AuthenticodeSignature -FilePath $TargetPath -ErrorAction SilentlyContinue } catch { $sig = $null }
  if ($sig -and $sig.Status -eq 'Valid') { return $false }
  Write-Host '[SelfSign] Creating/trusting a local code-signing certificate and signing the script...'
  $subject = "CN=VCFStretch Local Code Signing ($env:USERNAME@$env:COMPUTERNAME)"
  $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
    Where-Object { $_.Subject -like 'CN=VCFStretch Local Code Signing*' } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1
  if (-not $cert) {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject -CertStoreLocation 'Cert:\CurrentUser\My' -KeyAlgorithm RSA -KeyLength 3072 -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(5)
  }
  foreach ($store in 'Cert:\CurrentUser\Root','Cert:\CurrentUser\TrustedPublisher') {
    try { $null = $cert | Copy-Item -Destination $store -Force -ErrorAction SilentlyContinue } catch {}
  }
  $null = Set-AuthenticodeSignature -FilePath $TargetPath -Certificate $cert -ErrorAction Stop
  Write-Host '[SelfSign] Script signed.'
  return $true
}

try {
  $pwsh = $null
  $proc = Get-Process -Id $PID -ErrorAction SilentlyContinue
  if ($proc) { $pwsh = $proc.Path }
} catch { $pwsh = $null }
if (-not $pwsh) { $pwsh = 'pwsh.exe' }

if (-not $NoAutoSign -and -not $SignedOk) {
  try { $did = Ensure-SelfSigned -TargetPath $PSCommandPath } catch { $did = $false }
  & $pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "$PSCommandPath" -SignedOk -NoRelaunch
  exit $LASTEXITCODE
}

if (-not $NoRelaunch) {
  $ap = [Threading.Thread]::CurrentThread.ApartmentState
  if ($ap -ne 'STA') {
    & $pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "$PSCommandPath" -NoRelaunch -SignedOk
    exit $LASTEXITCODE
  }
}

# --- Run directory + logging ---
$script:ReportsBase = (Get-Location).Path
$script:RunDir = $null
$Global:LogFile = $null
$script:LogWarmupSync = 50
$script:logQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

function New-RunDir {
  param([string]$Base)
  if ([string]::IsNullOrWhiteSpace($Base) -or -not (Test-Path $Base)) { $Base = (Get-Location).Path }
  $d = Join-Path $Base ("VCFStretch-Run-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
  New-Item -ItemType Directory -Force -Path $d | Out-Null
  $Global:LogFile = Join-Path $d ("VCFStretch-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + ".log")
  '' | Out-File -FilePath $Global:LogFile -Encoding UTF8 -Force
  $script:RunDir = $d
  $d
}

function Write-Log {
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR')][string]$Level='INFO'
  )
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
  $line = "[$ts][$Level] $Message"
  try { if ($Global:LogFile) { Add-Content -Path $Global:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } } catch {}
  try {
    if ($script:txtLog -and $script:window) {
      if ($script:LogWarmupSync -gt 0) {
        $script:window.Dispatcher.Invoke([Action]{
          try { $script:txtLog.AppendText("$line`r`n"); $script:txtLog.ScrollToEnd() } catch {}
        },[System.Windows.Threading.DispatcherPriority]::Render)
        $script:LogWarmupSync--
      } else {
        $null = $script:window.Dispatcher.BeginInvoke([Action]{
          try { $script:txtLog.AppendText("$line`r`n"); $script:txtLog.ScrollToEnd() } catch {}
        })
      }
    } else {
      # UI log removed - skip enqueue
    }
  } catch {}
  Write-Host $line
}

function Get-HttpErrorDetail {
  param([object]$Ex)
  $code = ''
  $snippet = ''
  try {
    $resp = $Ex.Response
    if ($resp -and $resp -is [System.Net.Http.HttpResponseMessage]) {
      $code = [int]$resp.StatusCode
      try { $snippet = $resp.Content.ReadAsStringAsync().Result } catch { $snippet = $Ex.Message }
    }
  } catch {}
  if ($snippet -and $snippet.Length -gt 400) { $snippet = $snippet.Substring(0,400) }
  [pscustomobject]@{ Code=$code; Snippet=$snippet }
}

# --- Module handling (extends the foundation pattern) ---
function Has-Module([string]$Name){ !!(Get-Module -ListAvailable -Name $Name | Select-Object -First 1) }

function Ensure-Module {
  param([Parameter(Mandatory)][string]$Name)

  $ok = Has-Module $Name
  if ($ok) {
    try { Import-Module $Name -ErrorAction SilentlyContinue | Out-Null; return $true } catch { return $false }
  }

  try {
    $old = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue | Out-Null
    Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -AcceptLicense -ErrorAction Stop
    Import-Module $Name -ErrorAction SilentlyContinue | Out-Null
    Write-Log "$Name installed/updated and imported."
    return $true
  } catch {
    Write-Log "$Name install failed: $($_.Exception.Message)" 'ERROR'
    return $false
  } finally {
    $ProgressPreference = $old
  }
}

function Set-StatusText {
  param([System.Windows.Controls.TextBlock]$Label,[string]$Text,[string]$State)
  if (-not $Label) { return }
  $Label.Text = $Text
  switch ($State) {
    'OK'   { $Label.Foreground = [Windows.Media.Brushes]::LightGreen }
    'WARN' { $Label.Foreground = [Windows.Media.Brushes]::Gold }
    'FAIL' { $Label.Foreground = [Windows.Media.Brushes]::Tomato }
    default { $Label.Foreground = [Windows.Media.Brushes]::White }
  }
}

# --- SDDC Manager API helpers ---
function New-SddcToken {
  param(
    [Parameter(Mandatory)][string]$SddcHost,
    [Parameter(Mandatory)][string]$Username,
    [Parameter(Mandatory)][string]$Password
  )
  $base = "https://$SddcHost"
  $body = @{ username=$Username; password=$Password } | ConvertTo-Json
  try {
    $tok = Invoke-RestMethod -Method Post -Uri "$base/v1/tokens" -ContentType 'application/json' -Body $body -SkipCertificateCheck
    if (-not $tok.accessToken) { throw "Token response missing accessToken" }
    return [pscustomobject]@{ Base=$base; AccessToken=$tok.accessToken; RefreshToken=$tok.refreshToken }
  } catch {
    $d = Get-HttpErrorDetail -Ex $_.Exception
    throw "Token acquire failed. HTTP $($d.Code) $($d.Snippet)"
  }
}

function Invoke-SddcApi {
  param(
    [Parameter(Mandatory)][psobject]$Session,
    [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','PUT','DELETE')][string]$Method,
    [Parameter(Mandatory)][string]$Path,
    [object]$Body = $null,
    [switch]$Raw
  )
  $uri = "$($Session.Base)$Path"
  $headers = @{ Authorization = "Bearer $($Session.AccessToken)"; Accept='application/json' }
  try {
    if ($Method -in @('POST','PATCH','PUT')) {
      if ($Body -is [string]) {
        return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType 'application/json' -Body $Body -SkipCertificateCheck
      } elseif ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 20
        return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType 'application/json' -Body $json -SkipCertificateCheck
      } else {
        return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType 'application/json' -SkipCertificateCheck
      }
    } else {
      return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -SkipCertificateCheck
    }
  } catch {
    $d = Get-HttpErrorDetail -Ex $_.Exception
    throw "SDDC Manager API call failed: $Method $Path -> HTTP $($d.Code) $($d.Snippet)"
  }
}

function Get-Clusters {
  param([psobject]$Session)
  # Try a couple common response shapes (elements, clusters, etc.).
  $res = Invoke-SddcApi -Session $Session -Method GET -Path '/v1/clusters'
  if ($res.elements) { return @($res.elements) }
  if ($res.clusters) { return @($res.clusters) }
  return @($res)
}

function Get-Hosts {
  param([psobject]$Session)
  $res = Invoke-SddcApi -Session $Session -Method GET -Path '/v1/hosts'
  if ($res.elements) { return @($res.elements) }
  if ($res.hosts) { return @($res.hosts) }
  return @($res)
}

function Get-HostIdByFqdn {
  param([psobject]$Session,[Parameter(Mandatory)][string]$Fqdn)
  $all = Get-Hosts -Session $Session
  $m = $all | Where-Object { ($_.fqdn -eq $Fqdn) -or ($_.hostname -eq $Fqdn) -or ($_.FQDN -eq $Fqdn) } | Select-Object -First 1
  if ($m -and $m.id) { return [string]$m.id }
  # fallback: partial match
  $m = $all | Where-Object { (($_.fqdn + '') -like "$Fqdn*") } | Select-Object -First 1
  if ($m -and $m.id) { return [string]$m.id }
  return $null
}

function Start-ClusterNetworkQuery {
  param([psobject]$Session,[Parameter(Mandatory)][string]$ClusterId)

  $body = @{ name='VCENTER_NSXT_NETWORK_CONFIG' }
  try {
    $json = ($body | ConvertTo-Json)
    $respHeaders = @{}
    $res = Invoke-RestMethod -Method Post -Uri ("$($Session.Base)/v1/clusters/{0}/network/queries" -f $ClusterId) -Headers @{ Authorization = "Bearer $($Session.AccessToken)"; Accept='application/json' } -ContentType 'application/json' -Body $json -SkipCertificateCheck -ResponseHeadersVariable respHeaders

    if ($res -and $res.queryId) { return [string]$res.queryId }
    if ($res -and $res.id) { return [string]$res.id }

    try {
      $loc = $respHeaders['Location']
      if (-not $loc) { $loc = $respHeaders['location'] }
      if ($loc) {
        $locStr = ($loc | Select-Object -First 1)
        $parts = $locStr.Trim().Split('/')
        $qid = $parts[$parts.Count-1]
        if ($qid) { return [string]$qid }
      }
    } catch {}

    Write-Log "Network query did not return queryId/id; attempting to list queries..." 'WARN'
    try {
      $list = Invoke-RestMethod -Method Get -Uri ("$($Session.Base)/v1/clusters/{0}/network/queries" -f $ClusterId) -Headers @{ Authorization = "Bearer $($Session.AccessToken)"; Accept='application/json' } -SkipCertificateCheck
      if ($list -and $list.elements -and $list.elements.Count -gt 0) {
        $latest = $list.elements | Sort-Object -Property createdAt -Descending | Select-Object -First 1
        if ($latest -and $latest.id) { return [string]$latest.id }
        if ($latest -and $latest.queryId) { return [string]$latest.queryId }
      }
    } catch {}

    return $null
  } catch {
    $d = Get-HttpErrorDetail -Ex $_.Exception
    throw "Network query start failed. HTTP $($d.Code) $($d.Snippet)"
  }
}

function Get-ClusterNetworkQueryResult {
  param([psobject]$Session,[Parameter(Mandatory)][string]$ClusterId,[Parameter(Mandatory)][string]$QueryId)
  return Invoke-SddcApi -Session $Session -Method GET -Path ("/v1/clusters/{0}/network/queries/{1}" -f $ClusterId,$QueryId)
}



function Normalize-Arrow {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $s }
  # normalize HTML encoded and alternate arrows to standard '->'
  $x = $s -replace '&gt;','>'
  $x = $x -replace '→','->'
  $x = $x -replace '-\>','->'
  return $x
}

function Derive-Vds02FromVds01 {
  param([string]$Vds01)
  if ([string]::IsNullOrWhiteSpace($Vds01)) { return $null }
  # replace trailing 01 with 02 (assumption: naming always ends in 01/02)
  if ($Vds01 -match '^(.*)01$') { return ($Matches[1] + '02') }
  return $null
}

function Ensure-DefaultVmnicMapping {
  # Sets a safe default vmnic mapping when user selected 2 vDS and the field is empty or still using placeholders.
  try {
    if (-not $script:txtVmnicMapping -or -not $script:cmbVdsConfig) { return }
    $vdsText = ''
    try { if ($script:cmbVdsConfig.SelectedItem) { $vdsText = ($script:cmbVdsConfig.SelectedItem.Content + '').Trim() } } catch {}
    if ($vdsText -notlike '2*') { return }

    $cur = ($script:txtVmnicMapping.Text + '')
    $curNorm = (Normalize-Arrow $cur)

    $isBlank = [string]::IsNullOrWhiteSpace($curNorm.Trim())
    $isPlaceholder = ($curNorm -match 'VDS1/' ) -or ($curNorm -match 'vds01/' )

    if (-not $isBlank -and -not $isPlaceholder) { return }

    $vds01 = ''
    try { if ($script:txtNsxHostSwitchVds) { $vds01 = ($script:txtNsxHostSwitchVds.Text + '').Trim() } } catch {}
    if ([string]::IsNullOrWhiteSpace($vds01)) { $vds01 = 'vds01' }
    $vds02 = Derive-Vds02FromVds01 -Vds01 $vds01
    if ([string]::IsNullOrWhiteSpace($vds02)) { $vds02 = 'vds02' }

    $script:txtVmnicMapping.Text = "vmnic0->$vds01/uplink1; vmnic1->$vds01/uplink2; vmnic2->$vds02/uplink1; vmnic3->$vds02/uplink2"
    Write-Log "Defaulted vmnic mapping for 2 vDS: $($script:txtVmnicMapping.Text)"
  } catch {}
}


function Ensure-DefaultNetworkMappings {
  <#
    Deterministically fills / corrects the three error-prone network mapping fields:
      - Standard vmnic mapping (vmnicX -> vDS/uplinkY)
      - vDS uplink -> NSX uplink mapping (uplink1 -> uplink-1)
      - Active uplinks (NSX names: uplink-1,uplink-2)

    Assumptions:
      - NSX always uses 2 uplinks: uplink-1 and uplink-2
      - 1 vDS => 2 pNIC (vmnic0, vmnic1)
      - 2 vDS => 4 pNIC (vmnic0..3) with vds02 derived from vds01 via 01->02
      - Only vDS name changes; mappings are constant
  #>
  param(
    [switch]$Force
  )

  # Need a vDS name to derive anything useful
  $vds1 = ($script:txtNsxHostSwitchVds.Text + '').Trim()
  if ([string]::IsNullOrWhiteSpace($vds1)) { return }

  # Determine topology (1 vDS or 2 vDS)
  $vdsText  = ''
  try { if ($script:cmbVdsConfig.SelectedItem) { $vdsText = ($script:cmbVdsConfig.SelectedItem.Content + '').Trim() } } catch {}
  $vdsCount = if ($vdsText -like '2*') { 2 } else { 1 }

  # --- vDS uplink -> NSX uplink mapping (constant) ---
  $desiredMap = 'uplink1->uplink-1; uplink2->uplink-2'
  $curMap = (Normalize-Arrow ($script:txtVdsToNsxUplinkMap.Text + '')).Trim()
  if ($Force -or [string]::IsNullOrWhiteSpace($curMap)) {
    $script:txtVdsToNsxUplinkMap.Text = $desiredMap
    Write-Log "Auto-set vDS->NSX uplink mapping: $desiredMap"
  }

  # --- Active uplinks must be NSX names (uplink-1,uplink-2) ---
  $desiredActive = 'uplink-1,uplink-2'
  $curActive = (($script:txtActiveUplinks.Text + '')).Trim()
  if ($Force -or [string]::IsNullOrWhiteSpace($curActive)) {
    $script:txtActiveUplinks.Text = $desiredActive
    Write-Log "Auto-set Active Uplinks (NSX names): $desiredActive"
  } else {
    # Auto-correct common human error (uplink1,uplink2 -> uplink-1,uplink-2)
    $normalized = $curActive -replace '\buplink1\b','uplink-1' -replace '\buplink2\b','uplink-2'
    if ($normalized -ne $curActive) {
      $script:txtActiveUplinks.Text = $normalized
      Write-Log "Auto-corrected Active Uplinks to NSX names: $normalized"
    }
  }

  # --- Standard vmnic mapping ---
  $curVmnic = (Normalize-Arrow ($script:txtVmnicMapping.Text + '')).Trim()

  if ($vdsCount -eq 1) {
    $desiredVmnic = "vmnic0->$vds1/uplink1; vmnic1->$vds1/uplink2"
    if ($Force -or [string]::IsNullOrWhiteSpace($curVmnic)) {
      $script:txtVmnicMapping.Text = $desiredVmnic
      Write-Log "Auto-set vmnic mapping (1 vDS): $desiredVmnic"
    }
  } else {
    $vds2 = Derive-Vds02FromVds01 -Vds01 $vds1
    if ([string]::IsNullOrWhiteSpace($vds2)) { $vds2 = 'vds02' }
    $desiredVmnic = "vmnic0->$vds1/uplink1; vmnic1->$vds1/uplink2; vmnic2->$vds2/uplink1; vmnic3->$vds2/uplink2"
    if ($Force -or [string]::IsNullOrWhiteSpace($curVmnic)) {
      $script:txtVmnicMapping.Text = $desiredVmnic
      Write-Log "Auto-set vmnic mapping (2 vDS): $desiredVmnic"
    }
  }
}

# --- Input parsing helpers ---
function Parse-CommaList([string]$s){
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  return @($s.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Parse-IpRanges([string]$s){
  # expects "start-end,start-end"
  $ranges = @()
  foreach($r in (Parse-CommaList $s)) {
    if ($r -notmatch '-') { continue }
    $parts = $r.Split('-',2)
    $ranges += [pscustomobject]@{ start=$parts[0].Trim(); end=$parts[1].Trim() }
  }
  $ranges
}

function Parse-VmnicMapping([string]$s){
  # Format: "vmnic0->VDS1/uplink1; vmnic1->VDS1/uplink2; vmnic2->VDS2/uplink1; vmnic3->VDS2/uplink2"
  $out = @()
  foreach($chunk in ($s -split ';')) {
    $c = $chunk.Trim()
    if (-not $c) { continue }
    if ($c -notmatch '^(vmnic\d+)\s*->\s*([^/]+)/\s*(\S+)$') {
      continue
    }
    $vmnic = $Matches[1]
    $vds = $Matches[2].Trim()
    $uplink = $Matches[3].Trim()
    $out += [pscustomobject]@{ id=$vmnic; vdsName=$vds; uplink=$uplink }
  }
  $out
}

function Parse-UplinkMapping([string]$s){
  # Format: "uplink1->uplink-1" or "uplink1->uplink-1; uplink2->uplink-2"
  $out = @()
  foreach($chunk in ($s -split ';')) {
    $c = $chunk.Trim()
    if (-not $c) { continue }
    if ($c -notmatch '^(\S+)\s*->\s*(\S+)$') { continue }
    $out += [pscustomobject]@{ vdsUplinkName=$Matches[1].Trim(); nsxUplinkName=$Matches[2].Trim() }
  }
  $out
}

# --- Auto-populate UI from VCENTER_NSXT_NETWORK_CONFIG query output (best-effort, schema-tolerant) ---
function ConvertTo-FlatPairs {
 param(
  [Parameter(Mandatory)][object]$Obj,
  [string]$Prefix = ''
 )
 $pairs = [System.Collections.Generic.List[object]]::new()
 if ($null -eq $Obj) { return $pairs }

 if ($Obj -is [System.Collections.IDictionary] -or $Obj -is [pscustomobject]) {
  if ($Obj -is [System.Collections.IDictionary]) {
   foreach ($k in $Obj.Keys) {
    $v = $Obj[$k]
    $p = if ($Prefix) { "$Prefix.$k" } else { "$k" }
    foreach ($child in (ConvertTo-FlatPairs -Obj $v -Prefix $p)) { $pairs.Add($child) }
   }
  } else {
   foreach ($pr in $Obj.PSObject.Properties) {
    $k = $pr.Name
    $v = $pr.Value
    $p = if ($Prefix) { "$Prefix.$k" } else { "$k" }
    foreach ($child in (ConvertTo-FlatPairs -Obj $v -Prefix $p)) { $pairs.Add($child) }
   }
  }
  return $pairs
 }

 if ($Obj -is [System.Collections.IEnumerable] -and $Obj -isnot [string]) {
  $i = 0
  foreach ($item in $Obj) {
   $p = "{0}[{1}]" -f $Prefix,$i
   foreach ($child in (ConvertTo-FlatPairs -Obj $item -Prefix $p)) { $pairs.Add($child) }
   $i++
  }
  return $pairs
 }

 $pairs.Add([pscustomobject]@{ Path = $Prefix; Value = $Obj })
 return $pairs
}

function Find-FlatValue {
 param(
  [Parameter(Mandatory)][object[]]$FlatPairs,
  [Parameter(Mandatory)][string[]]$PathRegex,
  [ValidateSet('First','All')][string]$Mode = 'First'
 )
 $hits = foreach ($re in $PathRegex) {
  $FlatPairs | Where-Object { $_.Path -match $re } | Select-Object -ExpandProperty Value
 }
 $hits = @($hits) | Where-Object { $null -ne $_ -and ("" + $_).Trim() -ne '' }
 if ($Mode -eq 'All') { return $hits }
 return ($hits | Select-Object -First 1)
}

function Set-IfEmptyOrPrompt {
  param(
    [Parameter(Mandatory)][System.Windows.Controls.TextBox]$TextBox,
    [AllowNull()][AllowEmptyString()][string]$NewValue,
    [string]$Label = 'field',
    [switch]$Force
  )
  if (-not $TextBox) { return }
  if ([string]::IsNullOrWhiteSpace($NewValue)) { return }
  $current = ($TextBox.Text + '').Trim()
  if ($Force -or [string]::IsNullOrWhiteSpace($current)) {
    $TextBox.Text = $NewValue
    Write-Log "Auto-filled ${Label}: $NewValue"
  }
}

function AutoPopulate-FromNetworkQuery {
 param([Parameter(Mandatory)][object]$QueryResult)
 $flat = @(ConvertTo-FlatPairs -Obj $QueryResult)
 if (-not $flat -or $flat.Count -lt 1) {
  Write-Log "Auto-fill skipped: query result could not be flattened." 'WARN'
  return
 }

 $tepPoolName = Find-FlatValue -FlatPairs $flat -PathRegex @('ipAddressPoolsSpec\[\d+\]\.name$','ipAddressPool(s)?\[\d+\]\.name$','\.tep.*pool.*name$')
 $tepCidr = Find-FlatValue -FlatPairs $flat -PathRegex @('ipAddressPoolsSpec\[\d+\]\.subnets\[\d+\]\.cidr$','ipAddressPool.*subnets\[\d+\]\.cidr$','\.tep.*cidr$')
 $tepGw = Find-FlatValue -FlatPairs $flat -PathRegex @('ipAddressPoolsSpec\[\d+\]\.subnets\[\d+\]\.gateway$','ipAddressPool.*subnets\[\d+\]\.gateway$','\.tep.*gateway$')
 $tepStart = Find-FlatValue -FlatPairs $flat -PathRegex @('ipAddressPoolsSpec\[\d+\]\.subnets\[\d+\]\.ipAddressPoolRanges\[\d+\]\.start$','\.ip(Address)?PoolRanges\[\d+\]\.start$','\.tep.*range.*start$')
 $tepEnd = Find-FlatValue -FlatPairs $flat -PathRegex @('ipAddressPoolsSpec\[\d+\]\.subnets\[\d+\]\.ipAddressPoolRanges\[\d+\]\.end$','\.ip(Address)?PoolRanges\[\d+\]\.end$','\.tep.*range.*end$')

 $uplinkProfileName = Find-FlatValue -FlatPairs $flat -PathRegex @('uplinkProfiles\[\d+\]\.name$','uplinkProfile(s)?\[\d+\]\.name$','\.uplinkProfileName$')
 $transportVlan = Find-FlatValue -FlatPairs $flat -PathRegex @('uplinkProfiles\[\d+\]\.transportVlan$','\.transportVlan$')
 $teamPolicy = Find-FlatValue -FlatPairs $flat -PathRegex @('uplinkProfiles\[\d+\]\.teamings\[\d+\]\.policy$','\.teamings\[\d+\]\.policy$','\.teaming.*policy$')

 $activeUplinks = Find-FlatValue -FlatPairs $flat -PathRegex @('uplinkProfiles\[\d+\]\.teamings\[\d+\]\.activeUplinks\[\d+\]$') -Mode All
 $standbyUplinks = Find-FlatValue -FlatPairs $flat -PathRegex @('uplinkProfiles\[\d+\]\.teamings\[\d+\]\.standByUplinks\[\d+\]$') -Mode All

 $vdsName = Find-FlatValue -FlatPairs $flat -PathRegex @('nsxtHostSwitchConfigs\[\d+\]\.vdsName$','\.vdsName$')

 $pairVds = Find-FlatValue -FlatPairs $flat -PathRegex @('vdsUplinkToNsxUplink\[\d+\]\.vdsUplinkName$') -Mode All
 $pairNsx = Find-FlatValue -FlatPairs $flat -PathRegex @('vdsUplinkToNsxUplink\[\d+\]\.nsxUplinkName$') -Mode All
 $uplinkPairs = @()
 for ($i=0; $i -lt [Math]::Min($pairVds.Count,$pairNsx.Count); $i++) {
  $uplinkPairs += ("{0}->{1}" -f $pairVds[$i],$pairNsx[$i])
 }
 $uplinkPairsText = ($uplinkPairs -join '; ')

 Set-IfEmptyOrPrompt -TextBox $script:txtTepPoolName -NewValue ($tepPoolName + '') -Label 'TEP Pool Name'
 Set-IfEmptyOrPrompt -TextBox $script:txtTepCidr -NewValue ($tepCidr + '') -Label 'TEP CIDR'
 Set-IfEmptyOrPrompt -TextBox $script:txtTepGateway -NewValue ($tepGw + '') -Label 'TEP Gateway'
 Set-IfEmptyOrPrompt -TextBox $script:txtTepRangeStart -NewValue ($tepStart + '') -Label 'TEP Range Start'
 Set-IfEmptyOrPrompt -TextBox $script:txtTepRangeEnd -NewValue ($tepEnd + '') -Label 'TEP Range End'

 Set-IfEmptyOrPrompt -TextBox $script:txtUplinkProfileName -NewValue ($uplinkProfileName + '') -Label 'Uplink Profile Name'
 Set-IfEmptyOrPrompt -TextBox $script:txtTransportVlan -NewValue ($transportVlan + '') -Label 'Transport VLAN'
 Set-IfEmptyOrPrompt -TextBox $script:txtTeamingPolicy -NewValue ($teamPolicy + '') -Label 'Teaming Policy'

 if ($activeUplinks -and $activeUplinks.Count -gt 0) {
  Set-IfEmptyOrPrompt -TextBox $script:txtActiveUplinks -NewValue (($activeUplinks | ForEach-Object { ("$_").Trim() } | Where-Object { $_ }) -join ',') -Label 'Active Uplinks'
 }
 if ($standbyUplinks -and $standbyUplinks.Count -gt 0) {
  Set-IfEmptyOrPrompt -TextBox $script:txtStandbyUplinks -NewValue (($standbyUplinks | ForEach-Object { ("$_").Trim() } | Where-Object { $_ }) -join ',') -Label 'Standby Uplinks'
 }

 Set-IfEmptyOrPrompt -TextBox $script:txtNsxHostSwitchVds -NewValue ($vdsName + '') -Label 'NSX Host Switch vDS Name'
 if (-not [string]::IsNullOrWhiteSpace($uplinkPairsText)) {
  Set-IfEmptyOrPrompt -TextBox $script:txtVdsToNsxUplinkMap -NewValue $uplinkPairsText -Label 'vDS Uplink -> NSX Uplink map'
 }

 Write-Log 'Auto-fill completed (best-effort). Review fields before Generate/Validate/Execute.'
}

# --- Host picker: select UNASSIGNED/USEABLE hosts instead of manual paste ---
function Get-UnassignedUseableHosts {
 param([Parameter(Mandatory)][psobject]$Session)
 $hosts = @(Get-Hosts -Session $Session)
 $filtered = foreach ($h in $hosts) {
  $clusterAssigned = (-not [string]::IsNullOrWhiteSpace(($h.clusterId + ''))) -or ((-not ($null -eq $h.cluster)) -and (-not [string]::IsNullOrWhiteSpace(($h.cluster.id + ''))))
  $stateText = ((($h.status + ''),($h.state + ''),($h.hostState + ''),($h.assignmentState + ''),($h.usabilityStatus + ''),($h.usability + '')) -join ' | ')
  $looksGood = ($stateText -match 'UNASSIGNED[_-]?USEABLE') -or (($stateText -match 'UNASSIGNED') -and ($stateText -match 'USABLE|USEABLE|READY'))
  if (-not $clusterAssigned -and $looksGood) {
   [pscustomobject]@{ FQDN=(Coalesce $h.fqdn $h.hostname); Id=(Coalesce $h.id $h.hostId); Status=($h.status + ''); State=($h.state + ''); Details=$stateText }
  }
 }
 return @($filtered | Where-Object { -not [string]::IsNullOrWhiteSpace($_.FQDN) } | Sort-Object FQDN)
}


function Show-HostPicker {
 param([Parameter(Mandatory)][psobject]$Session)
 $data = Get-UnassignedUseableHosts -Session $Session
 $x = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select UNASSIGNED/USEABLE Hosts" Height="620" Width="980"
        WindowStartupLocation="CenterOwner" Background="#0f0f10" Foreground="#f3f3f3">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="Select one or more hosts (multi-select). Click OK to populate AZ2 host list." Margin="0,0,0,8"/>
    <DataGrid Grid.Row="1" x:Name="dg" AutoGenerateColumns="False" SelectionMode="Extended" SelectionUnit="FullRow" CanUserAddRows="False" IsReadOnly="True">
      <DataGrid.Columns>
        <DataGridTextColumn Header="FQDN" Binding="{Binding FQDN}" Width="2*"/>
        <DataGridTextColumn Header="Host ID" Binding="{Binding Id}" Width="2*"/>
        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="*"/>
        <DataGridTextColumn Header="State" Binding="{Binding State}" Width="*"/>
      </DataGrid.Columns>
    </DataGrid>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="btnCancel" Content="Cancel" Width="100" Margin="0,0,10,0"/>
      <Button x:Name="btnOk" Content="OK" Width="100"/>
    </StackPanel>
  </Grid>
</Window>
"@
 $w = [Windows.Markup.XamlReader]::Parse($x)
 $w.Owner = $script:window
 $dg = $w.FindName('dg')
 $btnOk = $w.FindName('btnOk')
 $btnCancel = $w.FindName('btnCancel')
 $dg.ItemsSource = $data
 # selection stored in $w.Tag
 $btnCancel.Add_Click({ $w.DialogResult = $false; $w.Close() })
 $btnOk.Add_Click({
  $w.Tag = @($dg.SelectedItems | ForEach-Object { $_.FQDN } | Where-Object { $_ })
  $w.DialogResult = $true
  $w.Close()
})
$null = $w.ShowDialog()
if ($w.Tag) { return @($w.Tag) }
return @()
}

# --- Optional PowerCLI vCenter verification: confirm AZ2 vmnic->vDS/uplink mapping matches expected ---
function Connect-VCenterIfNeeded {
 param(
  [Parameter(Mandatory)][string]$Server,
  [Parameter(Mandatory)][string]$User,
  [Parameter(Mandatory)][string]$Pass
 )

 $hasVcf = Has-Module 'VCF.PowerCLI'
 $hasCore = Has-Module 'VMware.VimAutomation.Core'
 if (-not $hasVcf -and -not $hasCore) {
  throw 'PowerCLI not found. Install VCF.PowerCLI (recommended) or VMware.PowerCLI (optional) from the UI.'
 }
 if ($hasVcf) { try { Import-Module VCF.PowerCLI -ErrorAction SilentlyContinue | Out-Null } catch {} }
 if ($hasCore) { try { Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue | Out-Null } catch {} }

 try { Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null } catch {}
 $existing = $global:DefaultVIServers | Where-Object { $_.Name -eq $Server -and $_.IsConnected } | Select-Object -First 1
 if ($existing) { return $existing }
 $sec = ConvertTo-SecureString $Pass -AsPlainText -Force
 $cred = [pscredential]::new($User,$sec)
 return (Connect-VIServer -Server $Server -Credential $cred -WarningAction SilentlyContinue -ErrorAction Stop)
}

# --- Build stretch spec (ESA/OSA, 1 or 2 vDS) ---
function New-ClusterStretchSpec {
  param(
    [Parameter(Mandatory)][psobject]$Session,
    [Parameter(Mandatory)][string]$ClusterId,
    [Parameter(Mandatory)][string]$StorageType,  # ESA or OSA
    [Parameter(Mandatory)][int]$VdsCount,         # 1 or 2
    [Parameter(Mandatory)][string]$Az1Name,
    [Parameter(Mandatory)][string]$Az2Name,
    [Parameter(Mandatory)][string[]]$Az2HostsFqdn,
    [Parameter(Mandatory)][string]$Az2NetworkProfileName,
    [Parameter(Mandatory)][bool]$Az2NetworkProfileIsDefault,
    [Parameter(Mandatory)][string]$VmnicMappingText,
    [Parameter(Mandatory)][string]$NsxHostSwitchVdsName,
    [Parameter(Mandatory)][string]$TepPoolName,
    [Parameter(Mandatory)][string]$TepCidr,
    [Parameter(Mandatory)][string]$TepGateway,
    [Parameter(Mandatory)][string]$TepRangeStart,
    [Parameter(Mandatory)][string]$TepRangeEnd,
    [Parameter(Mandatory)][string]$UplinkProfileName,
    [Parameter(Mandatory)][int]$TransportVlan,
    [Parameter(Mandatory)][string]$TeamingPolicy,
    [Parameter(Mandatory)][string]$ActiveUplinksCsv,
    [string]$StandbyUplinksCsv,
    [Parameter(Mandatory)][string]$VdsToNsxUplinkMapText,
    [bool]$DeployWithoutLicenseKeys,
    [bool]$IsEdgeClusterConfiguredForMultiAZ,
    [Parameter(Mandatory)][string]$WitnessFqdn,
    [Parameter(Mandatory)][string]$WitnessVsanIp,
    [Parameter(Mandatory)][string]$WitnessVsanCidr,
    [bool]$WitnessTrafficSharedWithVsanTraffic
  )

  $vmnics = Parse-VmnicMapping $VmnicMappingText
  if ($VdsCount -eq 1 -and $vmnics.Count -lt 2) { throw '1 vDS selection requires at least 2 vmnic mappings (vmnic0/vmnic1).' }
  if ($VdsCount -eq 2 -and $vmnics.Count -lt 4) { throw '2 vDS selection requires 4 vmnic mappings (vmnic0..vmnic3).' }

  $uplinkMap = Parse-UplinkMapping $VdsToNsxUplinkMapText
  if ($uplinkMap.Count -lt 2) { throw 'vDS uplink to NSX uplink mapping must include at least two pairs.' }

  $hostSpecs = @()
  foreach($fqdn in $Az2HostsFqdn) {
    $id = Get-HostIdByFqdn -Session $Session -Fqdn $fqdn
    if (-not $id) { throw "Could not resolve host ID for FQDN: $fqdn (ensure host is commissioned and visible in /v1/hosts)" }

    $hostSpecs += [pscustomobject]@{
      id = $id
      hostname = $fqdn
      azName = $Az2Name
      hostNetworkSpec = [pscustomobject]@{
        networkProfileName = $Az2NetworkProfileName
        vmNics = @($vmnics)
      }
    }
  }

  $ipPoolsSpec = @(
    [pscustomobject]@{
      name = $TepPoolName
      subnets = @(
        [pscustomobject]@{
          cidr = $TepCidr
          gateway = $TepGateway
          ipAddressPoolRanges = @(
            [pscustomobject]@{ start = $TepRangeStart; end = $TepRangeEnd }
          )
        }
      )
    }
  )

  $uplinkProfiles = @(
    [pscustomobject]@{
      name = $UplinkProfileName
      transportVlan = [int]$TransportVlan
      teamings = @(
        [pscustomobject]@{
          name = 'DEFAULT'
          policy = $TeamingPolicy
          standByUplinks = @(Parse-CommaList $StandbyUplinksCsv)
          activeUplinks = @(Parse-CommaList $ActiveUplinksCsv)
        }
      )
    }
  )

  $networkProfiles = @(
    [pscustomobject]@{
      isDefault = [bool]$Az2NetworkProfileIsDefault
      name = $Az2NetworkProfileName
      nsxtHostSwitchConfigs = @(
        [pscustomobject]@{
          ipAddressPoolName = $TepPoolName
          uplinkProfileName = $UplinkProfileName
          vdsName = $NsxHostSwitchVdsName
          vdsUplinkToNsxUplink = @($uplinkMap)
        }
      )
    }
  )

  $spec = [pscustomobject]@{
    clusterStretchSpec = [pscustomobject]@{
      deployWithoutLicenseKeys = [bool]$DeployWithoutLicenseKeys
      hostSpecs = @($hostSpecs)
      networkSpec = [pscustomobject]@{
        networkProfiles = @($networkProfiles)
        nsxClusterSpec = [pscustomobject]@{
          ipAddressPoolsSpec = @($ipPoolsSpec)
          uplinkProfiles = @($uplinkProfiles)
        }
      }
      isEdgeClusterConfiguredForMultiAZ = [bool]$IsEdgeClusterConfiguredForMultiAZ
      witnessSpec = [pscustomobject]@{
        fqdn = $WitnessFqdn
        vsanCidr = $WitnessVsanCidr
        vsanIp = $WitnessVsanIp
      }
      witnessTrafficSharedWithVsanTraffic = [bool]$WitnessTrafficSharedWithVsanTraffic
    }
  }

  return $spec
}

# --- Excel fill (template -> output) ---
function Write-FilledWorkbook {
  param(
    [Parameter(Mandatory)][string]$TemplatePath,
    [Parameter(Mandatory)][string]$SheetName,
    [Parameter(Mandatory)][hashtable]$ValueMap,
    [Parameter(Mandatory)][string]$OutPath
  )

  if (-not (Ensure-Module -Name 'ImportExcel')) { throw 'ImportExcel module not available.' }
  Copy-Item -Path $TemplatePath -Destination $OutPath -Force

  $pkg = Open-ExcelPackage -Path $OutPath
  try {
    $ws = $pkg.Workbook.Worksheets[$SheetName]
    if (-not $ws) { throw "Worksheet not found: $SheetName" }

    # Find header row where Required Item and Collected Data exist
    $headerRow = $null
    for ($r=1; $r -le 25; $r++) {
      $a = ($ws.Cells[$r,1].Text + '').Trim()
      $b = ($ws.Cells[$r,2].Text + '').Trim()
      $c = ($ws.Cells[$r,3].Text + '').Trim()
      if ($a -like '*Required Item*' -and $c -like '*Collected Data*') { $headerRow = $r; break }
    }
    if (-not $headerRow) { throw 'Could not locate header row in worksheet.' }

    # Determine column indexes by scanning header row
    $maxCol = $ws.Dimension.End.Column
    $colRequired = $null
    $colCollected = $null
    for ($c=1; $c -le $maxCol; $c++) {
      $txt = ($ws.Cells[$headerRow,$c].Text + '').Trim()
      if ($txt -like 'Required Item*') { $colRequired = $c }
      if ($txt -like 'Collected Data*') { $colCollected = $c }
    }
    if (-not $colRequired -or -not $colCollected) { throw 'Could not locate Required Item / Collected Data columns.' }

    # Fill rows
    $row = $headerRow + 1
    $filled = 0
    while ($row -le $ws.Dimension.End.Row) {
      $key = ($ws.Cells[$row,$colRequired].Text + '').Trim()
      if ([string]::IsNullOrWhiteSpace($key)) { break }

      if ($ValueMap.ContainsKey($key)) {
        $val = $ValueMap[$key]
        $ws.Cells[$row,$colCollected].Value = $val
        $filled += 1
      }
      $row += 1
    }

    Close-ExcelPackage $pkg
    return $filled
  } catch {
    try { Close-ExcelPackage $pkg } catch {}
    throw
  }
}

# --- UI (WPF) ---
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase -ErrorAction SilentlyContinue | Out-Null
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue | Out-Null

$xaml = @"

<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:sys="clr-namespace:System;assembly=mscorlib"
        Title="VCF 9 Stretched Cluster Automation (v{#VER#})"
        Height="900" Width="1600" MinHeight="760" MinWidth="1240"
        WindowStartupLocation="CenterScreen" Background="#0f0f10" Foreground="#f3f3f3">
  <Window.Resources>
    <SolidColorBrush x:Key="Bg" Color="#0f0f10"/>
    <SolidColorBrush x:Key="PanelBg" Color="#1c1c1e"/>
    <SolidColorBrush x:Key="Fg" Color="#f3f3f3"/>
    <SolidColorBrush x:Key="Border" Color="#3a3a3a"/>
    <SolidColorBrush x:Key="HeaderBg" Color="#2a2a2c"/>
 <SolidColorBrush x:Key="{x:Static SystemColors.HotTrackBrushKey}" Color="#f3f3f3"/>
 <SolidColorBrush x:Key="{x:Static SystemColors.ControlTextBrushKey}" Color="#f3f3f3"/>
 <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#1c1c1e"/>
  <!-- Override default Window/List brushes so ComboBox popups do not render as white-on-white -->
  <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="#1c1c1e"/>
  <SolidColorBrush x:Key="{x:Static SystemColors.WindowTextBrushKey}" Color="#f3f3f3"/>
  <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#3a3a3a"/>
  <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#f3f3f3"/>

    <Style TargetType="GroupBox"><Setter Property="Margin" Value="8"/><Setter Property="Padding" Value="8"/>
      <Setter Property="BorderBrush" Value="{StaticResource Border}"/><Setter Property="Foreground" Value="{StaticResource Fg}"/>
      <Setter Property="Background" Value="{StaticResource Bg}"/></Style>
    <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="Margin" Value="8,0,8,6"/></Style>
    <Style TargetType="CheckBox"><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="Margin" Value="8,4,8,4"/></Style>
    <Style TargetType="TextBox"><Setter Property="Margin" Value="8"/><Setter Property="Padding" Value="4"/><Setter Property="Height" Value="28"/>
      <Setter Property="Background" Value="{StaticResource PanelBg}"/><Setter Property="Foreground" Value="{StaticResource Fg}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Border}"/></Style>
    <Style TargetType="PasswordBox"><Setter Property="Margin" Value="8"/><Setter Property="Padding" Value="4"/><Setter Property="Height" Value="28"/>
      <Setter Property="Background" Value="{StaticResource PanelBg}"/><Setter Property="Foreground" Value="{StaticResource Fg}"/>
      <Setter Property="BorderBrush" Value="#565656"/></Style>
    <Style TargetType="ComboBox">
  <Setter Property="Margin" Value="8"/>
  <Setter Property="Padding" Value="6,3"/>
  <Setter Property="Height" Value="28"/>
  <Setter Property="Foreground" Value="{StaticResource Fg}"/>
  <Setter Property="Background" Value="{StaticResource PanelBg}"/>
  <Setter Property="BorderBrush" Value="{StaticResource Border}"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="ComboBox">
        <Grid>
          <ToggleButton x:Name="Toggle" Focusable="False" ClickMode="Press"
                        IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                        Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1">
            <ToggleButton.Template>
              <ControlTemplate TargetType="ToggleButton">
                <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="2">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition/>
                      <ColumnDefinition Width="22"/>
                    </Grid.ColumnDefinitions>
                    <ContentPresenter Grid.Column="0" Margin="6,3" VerticalAlignment="Center"
                                      Content="{Binding RelativeSource={RelativeSource TemplatedParent}, Path=TemplatedParent.SelectionBoxItem}"
                                      TextElement.Foreground="{StaticResource Fg}"/>
                    <Path Grid.Column="1" VerticalAlignment="Center" HorizontalAlignment="Center" Fill="{StaticResource Fg}"
                          Data="M 0 0 L 4 4 L 8 0 Z"/>
                  </Grid>
                </Border>
                <ControlTemplate.Triggers>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Bd" Property="BorderBrush" Value="#6a6a6a"/>
                  </Trigger>
                  <Trigger Property="IsChecked" Value="True">
                    <Setter TargetName="Bd" Property="BorderBrush" Value="#6a6a6a"/>
                  </Trigger>
                  <Trigger Property="IsEnabled" Value="False">
                    <Setter TargetName="Bd" Property="Opacity" Value="0.65"/>
                  </Trigger>
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </ToggleButton.Template>
          </ToggleButton>

          <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False">
            <Border Background="{StaticResource PanelBg}" BorderBrush="{StaticResource Border}" BorderThickness="1" CornerRadius="2">
              <ScrollViewer Margin="4" SnapsToDevicePixels="True">
                <ItemsPresenter/>
              </ScrollViewer>
            </Border>
          </Popup>
        </Grid>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<Style TargetType="ComboBoxItem">
  <Setter Property="Foreground" Value="{StaticResource Fg}"/>
  <Setter Property="Background" Value="{StaticResource PanelBg}"/>
  <Setter Property="Padding" Value="6,3"/>
  <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
  <Style.Triggers>
    <Trigger Property="IsHighlighted" Value="True">
      <Setter Property="Background" Value="#3a3a3a"/>
    </Trigger>
    <Trigger Property="IsSelected" Value="True">
      <Setter Property="Background" Value="#3a3a3a"/>
    </Trigger>
  </Style.Triggers>
</Style>
    <Style TargetType="Button"><Setter Property="Margin" Value="8,6,8,6"/><Setter Property="Padding" Value="8,4"/><Setter Property="Height" Value="28"/>
      <Setter Property="Background" Value="#2a2a2c"/><Setter Property="Foreground" Value="{StaticResource Fg}"/>
      <Setter Property="BorderBrush" Value="#565656"/></Style>
    <Style TargetType="DataGrid"><Setter Property="Margin" Value="8"/><Setter Property="Background" Value="{StaticResource PanelBg}"/>
      <Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="GridLinesVisibility" Value="All"/>
      <Setter Property="HeadersVisibility" Value="Column"/><Setter Property="BorderBrush" Value="{StaticResource Border}"/>
      <Setter Property="AlternationCount" Value="2"/><Setter Property="RowBackground" Value="#19191b"/>
      <Setter Property="AlternatingRowBackground" Value="#151517"/><Setter Property="HorizontalGridLinesBrush" Value="#303034"/>
      <Setter Property="VerticalGridLinesBrush" Value="#303034"/><Setter Property="SelectionUnit" Value="FullRow"/></Style>
    <Style TargetType="DataGridColumnHeader"><Setter Property="Foreground" Value="{StaticResource Fg}"/>
      <Setter Property="Background" Value="{StaticResource HeaderBg}"/><Setter Property="BorderBrush" Value="{StaticResource Border}"/>
      <Setter Property="FontWeight" Value="SemiBold"/></Style>
  </Window.Resources>

  <Grid Margin="8">
    
<Grid.RowDefinitions>
  <RowDefinition Height="Auto"/>
  <RowDefinition Height="*"/>
  <RowDefinition Height="Auto"/>
</Grid.RowDefinitions>

    <!-- Row 0: Prerequisites -->
    <GroupBox Header="Prerequisites" Grid.Row="0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="2*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0">
          <TextBlock x:Name="lblPS" Text="PowerShell: (checking...)" />
          <TextBlock x:Name="lblWPF" Text=".NET/WPF: (checking...)" />
          <TextBlock x:Name="lblImpExcel" Text="ImportExcel: (checking...)" />
          <TextBlock x:Name="lblPCLI" Text="VMware.PowerCLI (optional): (checking...)" />
<TextBlock x:Name="lblVCFPCLI" Text="VCF.PowerCLI (recommended): (checking...)" />
        </StackPanel>
        <StackPanel Grid.Column="1" VerticalAlignment="Center">
          <Button x:Name="btnRecheck" Content="Recheck" MinWidth="110"/>
        </StackPanel>
        <StackPanel Grid.Column="2" Orientation="Vertical" VerticalAlignment="Center">
          <Button x:Name="btnInstallImpExcel" Content="Install ImportExcel" MinWidth="170"/>
          <Button x:Name="btnInstallVCFPCLI" Content="Install VCF.PowerCLI" MinWidth="170"/>
<Button x:Name="btnInstallPCLI" Content="Install VMware.PowerCLI" MinWidth="170"/>
        </StackPanel>
      </Grid>
    </GroupBox>

    <!-- Row 1: Inputs -->
    
<Grid Grid.Row="1">
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="*"/>
    <ColumnDefinition Width="*"/>
    <ColumnDefinition Width="*"/>
  </Grid.ColumnDefinitions>

<ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
  <StackPanel><GroupBox Header="Config (Save/Load)" Margin="8">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="Save/load form values (excluding passwords) to speed up repeat runs." Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBox Grid.Column="0" x:Name="txtConfigPath" Height="28"/>
      <Button Grid.Column="1" x:Name="btnBrowseConfig" Content="Browse..." MinWidth="90"/>
      <Button Grid.Column="2" x:Name="btnLoadConfig" Content="Load" MinWidth="80"/>
      <Button Grid.Column="3" x:Name="btnSaveConfig" Content="Save" MinWidth="80"/>
    </Grid>
  </Grid>
</GroupBox>
<GroupBox Header="vCenter (required)" Margin="8">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="Connect to vCenter first. This enables Generate/Validate/Execute and provides inventory-based suggestions." Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
    <Grid Grid.Row="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="2*"/>
        <ColumnDefinition Width="2*"/>
        <ColumnDefinition Width="2*"/>
      </Grid.ColumnDefinitions>
      <TextBlock Grid.Row="0" Grid.Column="0" Text="vCenter FQDN" Margin="8,0,8,2"/>
      <TextBlock Grid.Row="0" Grid.Column="1" Text="Username" Margin="8,0,8,2"/>
      <TextBlock Grid.Row="0" Grid.Column="2" Text="Password" Margin="8,0,8,2"/>
      <TextBox Grid.Row="1" Grid.Column="0" x:Name="txtVCenterFqdn"/>
      <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtVCenterUser"/>
      <PasswordBox Grid.Row="1" Grid.Column="2" x:Name="pbVCenterPass"/>
    </Grid>
    <StackPanel Grid.Row="2" Orientation="Horizontal" VerticalAlignment="Center">
      <Button x:Name="btnVerifyVCenter" Content="Connect/Verify" MinWidth="140"/>
      <TextBlock x:Name="lblVCenterStatus" Text="Not connected" Margin="10,0,0,0" VerticalAlignment="Center" Foreground="#bfbfbf"/>
    </StackPanel>
  </Grid>
</GroupBox>
<GroupBox Header="SDDC Manager API (VCF 9 Fleet Management)" VerticalAlignment="Top">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="SDDC Manager FQDN" />
    <TextBox Grid.Row="1" x:Name="txtSddcHost" />
    <TextBlock Grid.Row="2" Text="Username" />
    <TextBlock Grid.Row="3" x:Name="txtSddcUserHint" Text="Use an SSO/AD account with VCF ADMIN role. Default: administrator@vsphere.local. (admin@local may also work.)" Margin="8,0,8,4" Foreground="#bfbfbf" FontSize="11" />
    <TextBox Grid.Row="4" x:Name="txtSddcUser" />
    <TextBlock Grid.Row="5" Text="Password" />
    <StackPanel Grid.Row="6" Orientation="Horizontal" VerticalAlignment="Center">
      <PasswordBox x:Name="pbSddcPass" Width="260" Margin="8"/>
      <Button x:Name="btnConnect" Content="Connect" MinWidth="120"/>
      <TextBlock x:Name="lblConnStatus" Text="Not connected" Margin="10,0,0,0" VerticalAlignment="Center" Foreground="#bfbfbf"/>
    </StackPanel>
  </Grid>
</GroupBox>
<GroupBox Header="Witness" VerticalAlignment="Top">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="Witness FQDN"/>
    <TextBox Grid.Row="1" x:Name="txtWitnessFqdn"/>
    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <StackPanel Grid.Column="0">
        <TextBlock Text="Witness vSAN IP"/>
        <TextBox x:Name="txtWitnessVsanIp"/>
      </StackPanel>
      <StackPanel Grid.Column="1">
        <TextBlock Text="Witness vSAN CIDR"/>
        <TextBox x:Name="txtWitnessVsanCidr"/>
      </StackPanel>
    </Grid>
    <CheckBox Grid.Row="3" x:Name="chkWitnessShared" Content="Witness traffic shared with vSAN traffic" IsChecked="False"/>
  </Grid>
</GroupBox>
      </StackPanel>
    </ScrollViewer>
<GroupBox Header="Stretch Options" Grid.Column="1">
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="Auto"/>
<RowDefinition Height="Auto"/>
<RowDefinition Height="Auto"/>
<RowDefinition Height="Auto"/>
<RowDefinition Height="18"/>
<RowDefinition Height="Auto"/>
<RowDefinition Height="Auto"/>
</Grid.RowDefinitions>

          <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="220"/>
              <ColumnDefinition Width="40"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="220"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Storage Type" VerticalAlignment="Center" Margin="8,0,8,0"/>
            <ComboBox Grid.Column="1" x:Name="cmbStorageType">
              <ComboBoxItem Content="ESA"/>
              <ComboBoxItem Content="OSA"/>
            </ComboBox>
            <TextBlock Grid.Column="3" Text="vDS" VerticalAlignment="Center" Margin="8,0,8,0"/>
            <ComboBox Grid.Column="4" x:Name="cmbVdsConfig">
              <ComboBoxItem Content="1 vDS"/>
              <ComboBoxItem Content="2 vDS"/>
            </ComboBox>
          </Grid>

          


          <Grid Grid.Row="1">
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="Auto"/>
    <ColumnDefinition Width="*"/>
  </Grid.ColumnDefinitions>
  <TextBlock Grid.Column="0" Text="Cluster" VerticalAlignment="Center" Margin="8,0,8,0"/>
  <ComboBox Grid.Column="1" x:Name="cmbCluster"/>
</Grid>

          <Grid Grid.Row="2">
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="Auto"/>
    <ColumnDefinition Width="*"/>
  </Grid.ColumnDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="Primary AZ (AZ1)" VerticalAlignment="Center" Margin="8,0,8,0"/>
  <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtAz1"/>
  <TextBlock Grid.Row="1" Grid.Column="0" Text="Secondary AZ (AZ2)" VerticalAlignment="Center" Margin="8,0,8,0"/>
  <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtAz2"/>
</Grid>
<TextBlock Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="5" Text="AZ1/AZ2 names are labels used in the stretch spec (azName). Use the same names as your vCenter/vSAN fault domain (primary/secondary AZ)." Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
<TextBlock Grid.Row="4" Text="" Margin="0"/>



          <GroupBox Grid.Row="5" Header="AZ2 Hosts (FQDNs, one per line)" Margin="8,6,8,6">
  <DockPanel LastChildFill="True">
    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,8">
      <Button x:Name="btnPickHosts" Content="Pick UNASSIGNED/USEABLE Hosts..." MinWidth="285" Height="30" VerticalAlignment="Center"/>
    </StackPanel>
    <TextBox x:Name="txtAz2Hosts" AcceptsReturn="True" Height="90" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
  </DockPanel>
</GroupBox>

          <StackPanel Grid.Row="6" Orientation="Horizontal" VerticalAlignment="Center">
            <CheckBox x:Name="chkDeployNoLic" Content="Deploy without license keys" IsChecked="True"/>
            <CheckBox x:Name="chkEdgeMultiAZ" Content="Edge cluster configured for Multi-AZ" IsChecked="False"/>
          </StackPanel>

        </Grid>
      </GroupBox>
<GroupBox Header="Stretch Spec Details" Grid.Column="2">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel>
<TextBlock Text="Tip: Click Collect to pull VCENTER_NSXT_NETWORK_CONFIG and auto-fill some fields." Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
            <TextBlock Text="AZ2 Network Profile" />
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="1*"/>
              </Grid.ColumnDefinitions>
              <TextBox Grid.Column="0" x:Name="txtAz2NetworkProfileName" />
              <CheckBox Grid.Column="0" x:Name="chkAz2NetworkProfileDefault" Content="isDefault" IsChecked="True"/>
            </Grid><TextBlock Text="Name Suffix" />
<TextBlock Text="Optional. Used for auto-generated names (e.g., -01, -02)." Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
<TextBox x:Name="txtNameSuffix" Text="01"/>


            <TextBlock Text="Standard vmnic mapping" />
<TextBlock Text="Format: vmnicX->vDSName/uplink. Example: vmnic0->VDS1/uplink1; vmnic1->VDS1/uplink1; vmnic2->VDS1/uplink1; vmnic3->VDS1/uplink1" Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
            <TextBox x:Name="txtVmnicMapping" Height="56" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" />
 <CheckBox x:Name="chkManualMapOverride" Content="Allow manual override of mapping fields" IsChecked="False"/>



            <TextBlock Text="NSX Host Switch vDS name" />
            <TextBox x:Name="txtNsxHostSwitchVds" />

            <TextBlock Text="vDS uplink to NSX uplink mapping" />
<TextBlock Text="Each pair is vDS uplink name -> NSX uplink name. Example: uplink1->uplink1; uplink2->uplink2" Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
<TextBlock Text="Example: uplink1->uplink-1; uplink2->uplink-2" Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
            <TextBox x:Name="txtVdsToNsxUplinkMap" />

            <TextBlock Text="NSX IP Address Pool (ipAddressPoolsSpec)" />
<TextBlock Text="Required. Name must match nsxtHostSwitchConfigs.ipAddressPoolName. Auto-generated using suffix if blank." Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="1*"/>
    <ColumnDefinition Width="1*"/>
    <ColumnDefinition Width="1*"/>
  </Grid.ColumnDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="Pool Name (auto)" Margin="8,0,8,2"/>
  <TextBlock Grid.Row="0" Grid.Column="1" Text="CIDR" Margin="8,0,8,2"/>
  <TextBlock Grid.Row="0" Grid.Column="2" Text="Gateway" Margin="8,0,8,2"/>
  <TextBox Grid.Row="1" Grid.Column="0" x:Name="txtTepPoolName" Visibility="Collapsed"/>
  <TextBlock Grid.Row="1" Grid.Column="0" x:Name="lblDerivedIpPoolName" Text="(auto)" Margin="8"/>
  <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtTepCidr"/>
  <TextBox Grid.Row="1" Grid.Column="2" x:Name="txtTepGateway"/>
</Grid>
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="1*"/>
    <ColumnDefinition Width="1*"/>
  </Grid.ColumnDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="IP Range Start" Margin="8,0,8,2"/>
  <TextBlock Grid.Row="0" Grid.Column="1" Text="IP Range End" Margin="8,0,8,2"/>
  <TextBox Grid.Row="1" Grid.Column="0" x:Name="txtTepRangeStart"/>
  <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtTepRangeEnd"/>
</Grid>
<TextBlock Text="NSX Uplink Profile (uplinkProfiles)" />
<TextBlock Text="Required. Name must match nsxtHostSwitchConfigs.uplinkProfileName. Auto-generated using suffix if blank." Foreground="#bfbfbf" FontSize="11" Margin="8,0,8,6"/>
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="1.2*"/>
    <ColumnDefinition Width="0.8*"/>
    <ColumnDefinition Width="1*"/>
  </Grid.ColumnDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="Profile Name (auto)" Margin="8,0,8,2"/>
  <TextBlock Grid.Row="0" Grid.Column="1" Text="Transport VLAN" Margin="8,0,8,2"/>
  <TextBlock Grid.Row="0" Grid.Column="2" Text="Teaming Policy" Margin="8,0,8,2"/>
  <TextBox Grid.Row="1" Grid.Column="0" x:Name="txtUplinkProfileName" Visibility="Collapsed"/>
  <TextBlock Grid.Row="1" Grid.Column="0" x:Name="lblDerivedUplinkProfileName" Text="(auto)" Margin="8"/>
  <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtTransportVlan"/>
  <TextBox Grid.Row="1" Grid.Column="2" x:Name="txtTeamingPolicy"/>
</Grid>
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="1*"/>
    <ColumnDefinition Width="1*"/>
  </Grid.ColumnDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="Active Uplinks (CSV)" Margin="8,0,8,2"/>
  <TextBlock Grid.Row="0" Grid.Column="1" Text="Standby Uplinks (CSV)" Margin="8,0,8,2"/>
  <TextBox Grid.Row="1" Grid.Column="0" x:Name="txtActiveUplinks"/>
  <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtStandbyUplinks"/>
</Grid>
</StackPanel>
        </ScrollViewer>
      </GroupBox>
  </Grid>

      <!-- Row 2: Actions -->

    <GroupBox Header="Actions" Grid.Row="2">
      <Grid Margin="8">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Grid.Column="0">
          <TextBlock Text="Reports Path:" Margin="0,0,8,0" VerticalAlignment="Center"/>
          <TextBox x:Name="txtReportsPath" MinWidth="520" IsReadOnly="True" Height="28"/>
          <Button x:Name="btnBrowseReports" Content="Browse..." Margin="8,0,0,0" MinWidth="110"/>
        </StackPanel>

        <UniformGrid Grid.Column="1" Rows="1" Columns="7" Margin="12,0,0,0" HorizontalAlignment="Right">
          <Button x:Name="btnOpenOut" Content="Open Reports" MinWidth="120"/>
<Button x:Name="btnOpenLogFolder" Content="Open Log Location" MinWidth="140"/>
          <Button x:Name="btnCollect" Content="Collect" MinWidth="90"/>
          <Button x:Name="btnGenerate" Content="Generate" MinWidth="90" IsEnabled="False"/>
          <Button x:Name="btnValidate" Content="Validate" MinWidth="90" IsEnabled="False"/>
          <Button x:Name="btnExecute" Content="Execute" MinWidth="90" IsEnabled="False"/>
          <Button x:Name="btnClose" Content="Close" MinWidth="90"/>
        </UniformGrid>
      </Grid>
    </GroupBox>

  </Grid>
</Window>

"@


$xaml = $xaml.Replace('{#VER#}', $Global:VCFStretchVersion)

try {
  $xaml = $xaml.Replace('\"','"')
$script:window = [Windows.Markup.XamlReader]::Parse($xaml)
} catch {
  [System.Windows.MessageBox]::Show("XAML parse failed:`r`n$($_.Exception.Message)","XAML Error",'OK','Error') | Out-Null
  throw
}

$script:window.WindowState = 'Maximized'

# --- Find controls ---
$script:txtReports = $script:window.FindName('txtReportsPath')
$script:btnBrowseReports = $script:window.FindName('btnBrowseReports')
$script:btnOpenOut = $script:window.FindName('btnOpenOut')
$script:btnOpenLogFolder = $script:window.FindName('btnOpenLogFolder')
$script:btnClose = $script:window.FindName('btnClose')
$script:btnRecheck = $script:window.FindName('btnRecheck')
$script:btnInstallImpExcel = $script:window.FindName('btnInstallImpExcel')
$script:btnInstallVCFPCLI = $script:window.FindName('btnInstallVCFPCLI')
$script:btnInstallPCLI = $script:window.FindName('btnInstallPCLI')
$script:lblPS = $script:window.FindName('lblPS')
$script:lblWPF = $script:window.FindName('lblWPF')
$script:lblImpExcel = $script:window.FindName('lblImpExcel')
$script:lblPCLI = $script:window.FindName('lblPCLI')
$script:lblVCFPCLI = $script:window.FindName('lblVCFPCLI')
$script:txtSddcHost = $script:window.FindName('txtSddcHost')
$script:txtSddcUser = $script:window.FindName('txtSddcUser')
$script:pbSddcPass = $script:window.FindName('pbSddcPass')
$script:btnConnect = $script:window.FindName('btnConnect')
$script:lblConnStatus = $script:window.FindName('lblConnStatus')
$script:cmbStorageType = $script:window.FindName('cmbStorageType')
$script:cmbVdsConfig = $script:window.FindName('cmbVdsConfig')
$script:cmbCluster = $script:window.FindName('cmbCluster')
$script:txtAz1 = $script:window.FindName('txtAz1')
$script:txtAz2 = $script:window.FindName('txtAz2')
$script:txtAz2Hosts = $script:window.FindName('txtAz2Hosts')
$script:btnPickHosts = $script:window.FindName('btnPickHosts')
$script:btnVerifyVCenter = $script:window.FindName('btnVerifyVCenter')
$script:lblVCenterStatus = $script:window.FindName('lblVCenterStatus')
$script:chkRequireVCenterVerify = $script:window.FindName('chkRequireVCenterVerify')
$script:pbVCenterPass = $script:window.FindName('pbVCenterPass')
$script:txtVCenterUser = $script:window.FindName('txtVCenterUser')
$script:txtVCenterFqdn = $script:window.FindName('txtVCenterFqdn')
$script:chkDeployNoLic = $script:window.FindName('chkDeployNoLic')
$script:chkEdgeMultiAZ = $script:window.FindName('chkEdgeMultiAZ')
$script:btnCollect = $script:window.FindName('btnCollect')
$script:btnGenerate = $script:window.FindName('btnGenerate')
$script:btnValidate = $script:window.FindName('btnValidate')
$script:btnExecute = $script:window.FindName('btnExecute')

$script:txtAz2NetworkProfileName = $script:window.FindName('txtAz2NetworkProfileName')
$script:chkAz2NetworkProfileDefault = $script:window.FindName('chkAz2NetworkProfileDefault')
$script:txtVmnicMapping = $script:window.FindName('txtVmnicMapping')
$script:txtNsxHostSwitchVds = $script:window.FindName('txtNsxHostSwitchVds')
$script:txtVdsToNsxUplinkMap = $script:window.FindName('txtVdsToNsxUplinkMap')
$script:txtTepPoolName = $script:window.FindName('txtTepPoolName')
$script:txtTepCidr = $script:window.FindName('txtTepCidr')
$script:txtTepGateway = $script:window.FindName('txtTepGateway')
$script:txtTepRangeStart = $script:window.FindName('txtTepRangeStart')
$script:txtTepRangeEnd = $script:window.FindName('txtTepRangeEnd')
$script:txtUplinkProfileName = $script:window.FindName('txtUplinkProfileName')
$script:txtTransportVlan = $script:window.FindName('txtTransportVlan')
$script:txtTeamingPolicy = $script:window.FindName('txtTeamingPolicy')
$script:txtActiveUplinks = $script:window.FindName('txtActiveUplinks')
$script:txtStandbyUplinks = $script:window.FindName('txtStandbyUplinks')
$script:txtWitnessFqdn = $script:window.FindName('txtWitnessFqdn')
$script:txtWitnessVsanIp = $script:window.FindName('txtWitnessVsanIp')
$script:txtWitnessVsanCidr = $script:window.FindName('txtWitnessVsanCidr')
$script:chkWitnessShared = $script:window.FindName('chkWitnessShared')
$script:txtConfigPath = $script:window.FindName('txtConfigPath')
$script:btnBrowseConfig = $script:window.FindName('btnBrowseConfig')
$script:btnLoadConfig = $script:window.FindName('btnLoadConfig')
$script:btnSaveConfig = $script:window.FindName('btnSaveConfig')
$script:txtNameSuffix = $script:window.FindName('txtNameSuffix')
$script:lblDerivedIpPoolName = $script:window.FindName('lblDerivedIpPoolName')
$script:lblDerivedUplinkProfileName = $script:window.FindName('lblDerivedUplinkProfileName')
$script:chkManualMapOverride = $script:window.FindName('chkManualMapOverride')


function Set-MappingFieldsReadOnly {
  # Read-only by default; enable editing only when override box is checked.
  $allow = $false
  try { $allow = [bool]$script:chkManualMapOverride.IsChecked } catch { $allow = $false }
  $ro = -not $allow

  try { if ($script:txtVmnicMapping) { $script:txtVmnicMapping.IsReadOnly = $ro } } catch {}
  try { if ($script:txtVdsToNsxUplinkMap) { $script:txtVdsToNsxUplinkMap.IsReadOnly = $ro } } catch {}
  try { if ($script:txtActiveUplinks) { $script:txtActiveUplinks.IsReadOnly = $ro } } catch {}
  try { if ($script:txtStandbyUplinks) { $script:txtStandbyUplinks.IsReadOnly = $ro } } catch {}

  if (-not $allow) {
    Ensure-DefaultNetworkMappings -Force
  }
}


# --- UI log pump ---
$script:uiTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:uiTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$script:uiTimer.add_Tick({
  try {
    $sb = New-Object System.Text.StringBuilder
    while ($true) {
      if (-not $script:logQueue.TryDequeue([ref]$line)) { break }
      [void]$sb.Append($line)
    }
    if ($sb.Length -gt 0 -and $script:txtLog) {
      $script:txtLog.AppendText($sb.ToString())
      $script:txtLog.ScrollToEnd()
    }
  } catch {}
})

# --- State ---
$script:SddcSession = $null
$script:SelectedSheetName = $null
$script:ClusterMap = @{}   # display -> id
$script:LastNetworkQueryResult = $null # last VCENTER_NSXT_NETWORK_CONFIG query output
$script:VCenterVerified = $false
$script:VCenterServer = $null


function Prereq-Check {
  $ok = $true
  $isPS7 = $PSVersionTable.PSVersion.Major -ge 7
  Set-StatusText -Label $script:lblPS -Text ("PowerShell {0}" -f $PSVersionTable.PSVersion) -State $(if($isPS7){'OK'}else{'FAIL'})
  $ok = $ok -and $isPS7

  Set-StatusText -Label $script:lblWPF -Text '.NET/WPF: OK' -State 'OK'

  $hasImp = Has-Module 'ImportExcel'
  Set-StatusText -Label $script:lblImpExcel -Text ($hasImp ? 'ImportExcel: Found' : 'ImportExcel: Not found') -State ($hasImp ? 'OK':'WARN')

  $hasPcli = Has-Module 'VMware.VimAutomation.Core'
 Set-StatusText -Label $script:lblPCLI -Text ($hasPcli ? 'VMware.PowerCLI: Found' : 'VMware.PowerCLI: Not found (optional)') -State ($hasPcli ? 'OK':'WARN')
 $hasVcfPcli = Has-Module 'VCF.PowerCLI'
 Set-StatusText -Label $script:lblVCFPCLI -Text ($hasVcfPcli ? 'VCF.PowerCLI: Found' : 'VCF.PowerCLI: Not found (recommended)') -State ($hasVcfPcli ? 'OK':'WARN')
 return $ok
}

function Get-SelectedTopology {
  $storage = ($script:cmbStorageType.SelectedItem.Content + '').Trim()
  $vdsText = ($script:cmbVdsConfig.SelectedItem.Content + '').Trim()
  $vdsCount = if ($vdsText -like '2*') { 2 } else { 1 }

  $sheet = switch ("$storage-$vdsCount") {
    'ESA-1' { 'ESA_1vDS_2NIC' }
    'ESA-2' { 'ESA_2vDS_4NIC' }
    'OSA-1' { 'OSA_1vDS_2NIC' }
 'OSA-2' { 'OSA_2vDS_4NIC' }
 default { $null }
  }
  [pscustomobject]@{ StorageType=$storage; VdsCount=$vdsCount; SheetName=$sheet; IsSupported = -not [string]::IsNullOrWhiteSpace($sheet) }
}

function Refresh-TopologyUI {
  $t = Get-SelectedTopology
  $script:SelectedSheetName = $t.SheetName
  $script:btnGenerate.IsEnabled = ($t.IsSupported -and $script:SddcSession -ne $null -and $script:VCenterVerified)
  $script:btnValidate.IsEnabled = $script:btnGenerate.IsEnabled
  $script:btnExecute.IsEnabled = $script:btnGenerate.IsEnabled
 $script:btnPickHosts.IsEnabled = ($script:SddcSession -ne $null)
  if (-not $t.IsSupported) { Write-Log "Topology selection not supported by template (Storage=$($t.StorageType), vDS=$($t.VdsCount))." 'WARN' }
}

function Get-Az2HostsFromUi {
  $raw = ($script:txtAz2Hosts.Text + '')
  $hosts = @($raw -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
  return $hosts
}

function Get-SelectedClusterId {
  $sel = $script:cmbCluster.SelectedItem
  if (-not $sel) { return $null }
  $label = $sel.ToString()
  if ($script:ClusterMap.ContainsKey($label)) { return $script:ClusterMap[$label] }
  return $null
}

function Build-ValueMapForSheet {
  param([pscustomobject]$Spec,[string]$ClusterName,[string]$ClusterId,[string]$Az1,[string]$Az2,[string[]]$Az2Hosts)

  # Map workbook "Required Item" labels -> values.
  # Note: Labels must match the exact text in the template. If the template changes, update these keys.
  $m = @{}

  $storage = ($script:cmbStorageType.SelectedItem.Content + '').Trim()
  $m['Cluster storage type (ESA/OSA)'] = $storage
  $m['Management cluster name'] = $ClusterName
  $m['Management cluster ID (UUID)'] = $ClusterId
  $m['Primary AZ name (AZ1)'] = $Az1
  $m['Secondary AZ name (AZ2)'] = $Az2
  $m['Edge cluster configured for Multi-AZ (true/false)'] = (([bool]$script:chkEdgeMultiAZ.IsChecked).ToString().ToLower())
  $m['Deploy without license keys (true/false)'] = (([bool]$script:chkDeployNoLic.IsChecked).ToString().ToLower())

  # Hosts (up to 4 as template expects)
  for ($i=0; $i -lt [Math]::Min(4,$Az2Hosts.Count); $i++) {
    $idx = $i + 1
    $fqdn = $Az2Hosts[$i]
    $id = $Spec.clusterStretchSpec.hostSpecs[$i].id
    $m["Host $idx – SDDC Manager host ID (UUID)"] = $id
    $m["Host $idx – FQDN"] = $fqdn
  }

  # vmnic mapping + vDS names
  $m['Standard vmnic mapping (applies to all AZ2 hosts)'] = ($script:txtVmnicMapping.Text + '')
  # Optional vDS lines differ per sheet; populate if present.
  $m['vDS #1 name (NSX/Management/VM traffic)'] = ($script:txtNsxHostSwitchVds.Text + '')
  $m['vDS #1 name'] = ($script:txtNsxHostSwitchVds.Text + '')

  # Network profile
  $m['AZ2 Network Profile name (VCF/SDDC Manager)'] = ($script:txtAz2NetworkProfileName.Text + '')
  $m['AZ2 Network Profile isDefault (true/false)'] = (([bool]$script:chkAz2NetworkProfileDefault.IsChecked).ToString().ToLower())

  # NSX TEP pool details
  $m['NSX Host Switch – vDS name'] = ($script:txtNsxHostSwitchVds.Text + '')
  $m['NSX Host Switch – AZ2 uplink profile name'] = ($script:txtUplinkProfileName.Text + '')
  $m['NSX Host Switch – AZ2 TEP IP pool name'] = ($script:txtTepPoolName.Text + '')
  $m['vDS uplink to NSX uplink mapping – pair 1'] = ((Parse-UplinkMapping ($script:txtVdsToNsxUplinkMap.Text + '') | Select-Object -First 1 | ForEach-Object { "$($_.vdsUplinkName)->$($_.nsxUplinkName)" }) -join '')
  $m['vDS uplink to NSX uplink mapping – pair 2'] = ((Parse-UplinkMapping ($script:txtVdsToNsxUplinkMap.Text + '') | Select-Object -Skip 1 -First 1 | ForEach-Object { "$($_.vdsUplinkName)->$($_.nsxUplinkName)" }) -join '')
  $m['AZ2 TEP IP pool: CIDR'] = ($script:txtTepCidr.Text + '')
  $m['AZ2 TEP IP pool: Gateway'] = ($script:txtTepGateway.Text + '')
  $m['AZ2 TEP IP pool: IP range start'] = ($script:txtTepRangeStart.Text + '')
  $m['AZ2 TEP IP pool: IP range end'] = ($script:txtTepRangeEnd.Text + '')

  # Uplink profile
  $m['AZ2 Uplink profile: Transport VLAN'] = ($script:txtTransportVlan.Text + '')
  $m['AZ2 Uplink profile: Teaming policy'] = ($script:txtTeamingPolicy.Text + '')
  $m['AZ2 Uplink profile: Active uplinks'] = ($script:txtActiveUplinks.Text + '')
  $m['AZ2 Uplink profile: Standby uplinks (if any)'] = ($script:txtStandbyUplinks.Text + '')

  # Witness
  $m['Witness FQDN'] = ($script:txtWitnessFqdn.Text + '')
  $m['Witness vSAN IP'] = ($script:txtWitnessVsanIp.Text + '')
  $m['Witness vSAN CIDR'] = ($script:txtWitnessVsanCidr.Text + '')
  $m['Witness traffic shared with vSAN traffic (true/false)'] = (([bool]$script:chkWitnessShared.IsChecked).ToString().ToLower())

  return $m
}



function Get-UiConfig {
  return [pscustomobject]@{
    Version = $Global:VCFStretchVersion
    VCenterFqdn = ($script:txtVCenterFqdn.Text + '').Trim()
    VCenterUser = ($script:txtVCenterUser.Text + '').Trim()
    SddcHost = ($script:txtSddcHost.Text + '').Trim()
    SddcUser = ($script:txtSddcUser.Text + '').Trim()
    ClusterLabel = if ($script:cmbCluster.SelectedItem) { $script:cmbCluster.SelectedItem.ToString() } else { '' }
    StorageType = if ($script:cmbStorageType.SelectedItem) { ($script:cmbStorageType.SelectedItem.Content + '').Trim() } else { '' }
    VdsConfig = if ($script:cmbVdsConfig.SelectedItem) { ($script:cmbVdsConfig.SelectedItem.Content + '').Trim() } else { '' }
    Az1 = ($script:txtAz1.Text + '').Trim()
    Az2 = ($script:txtAz2.Text + '').Trim()
    Az2Hosts = ($script:txtAz2Hosts.Text + '')
    DeployNoLic = [bool]$script:chkDeployNoLic.IsChecked
    EdgeMultiAz = [bool]$script:chkEdgeMultiAZ.IsChecked
    Az2NetworkProfileName = ($script:txtAz2NetworkProfileName.Text + '').Trim()
    Az2NetworkProfileDefault = [bool]$script:chkAz2NetworkProfileDefault.IsChecked
    NameSuffix = ($script:txtNameSuffix.Text + '').Trim()
    VmnicMapping = ($script:txtVmnicMapping.Text + '').Trim()
    NsxHostSwitchVds = ($script:txtNsxHostSwitchVds.Text + '').Trim()
    VdsToNsxUplinkMap = ($script:txtVdsToNsxUplinkMap.Text + '').Trim()
    TepCidr = ($script:txtTepCidr.Text + '').Trim()
    TepGateway = ($script:txtTepGateway.Text + '').Trim()
    TepRangeStart = ($script:txtTepRangeStart.Text + '').Trim()
    TepRangeEnd = ($script:txtTepRangeEnd.Text + '').Trim()
    TransportVlan = ($script:txtTransportVlan.Text + '').Trim()
    TeamingPolicy = ($script:txtTeamingPolicy.Text + '').Trim()
    ActiveUplinks = ($script:txtActiveUplinks.Text + '').Trim()
    StandbyUplinks = ($script:txtStandbyUplinks.Text + '').Trim()
    WitnessFqdn = ($script:txtWitnessFqdn.Text + '').Trim()
    WitnessVsanIp = ($script:txtWitnessVsanIp.Text + '').Trim()
    WitnessVsanCidr = ($script:txtWitnessVsanCidr.Text + '').Trim()
    WitnessShared = [bool]$script:chkWitnessShared.IsChecked
  }
}

function Apply-UiConfig { param([Parameter(Mandatory)]$Cfg)
  if ($Cfg.VCenterFqdn) { $script:txtVCenterFqdn.Text = $Cfg.VCenterFqdn }
  if ($Cfg.VCenterUser) { $script:txtVCenterUser.Text = $Cfg.VCenterUser }
  if ($Cfg.SddcHost) { $script:txtSddcHost.Text = $Cfg.SddcHost }
  if ($Cfg.SddcUser) { $script:txtSddcUser.Text = $Cfg.SddcUser }
  if ($Cfg.Az1) { $script:txtAz1.Text = $Cfg.Az1 }
  if ($Cfg.Az2) { $script:txtAz2.Text = $Cfg.Az2 }
  if ($null -ne $Cfg.Az2Hosts) { $script:txtAz2Hosts.Text = $Cfg.Az2Hosts }
  if ($Cfg.Az2NetworkProfileName) { $script:txtAz2NetworkProfileName.Text = $Cfg.Az2NetworkProfileName }
  if ($null -ne $Cfg.Az2NetworkProfileDefault) { $script:chkAz2NetworkProfileDefault.IsChecked = [bool]$Cfg.Az2NetworkProfileDefault }
  if ($Cfg.NameSuffix) { $script:txtNameSuffix.Text = $Cfg.NameSuffix }
  if ($Cfg.VmnicMapping) { $script:txtVmnicMapping.Text = $Cfg.VmnicMapping }
  if ($Cfg.NsxHostSwitchVds) { $script:txtNsxHostSwitchVds.Text = $Cfg.NsxHostSwitchVds }
  if ($Cfg.VdsToNsxUplinkMap) { $script:txtVdsToNsxUplinkMap.Text = $Cfg.VdsToNsxUplinkMap }
  if ($Cfg.TepCidr) { $script:txtTepCidr.Text = $Cfg.TepCidr }
  if ($Cfg.TepGateway) { $script:txtTepGateway.Text = $Cfg.TepGateway }
  if ($Cfg.TepRangeStart) { $script:txtTepRangeStart.Text = $Cfg.TepRangeStart }
  if ($Cfg.TepRangeEnd) { $script:txtTepRangeEnd.Text = $Cfg.TepRangeEnd }
  if ($Cfg.TransportVlan) { $script:txtTransportVlan.Text = $Cfg.TransportVlan }
  if ($Cfg.TeamingPolicy) { $script:txtTeamingPolicy.Text = $Cfg.TeamingPolicy }
  if ($Cfg.ActiveUplinks) { $script:txtActiveUplinks.Text = $Cfg.ActiveUplinks }
  if ($null -ne $Cfg.StandbyUplinks) { $script:txtStandbyUplinks.Text = $Cfg.StandbyUplinks }
  if ($Cfg.WitnessFqdn) { $script:txtWitnessFqdn.Text = $Cfg.WitnessFqdn }
  if ($Cfg.WitnessVsanIp) { $script:txtWitnessVsanIp.Text = $Cfg.WitnessVsanIp }
  if ($Cfg.WitnessVsanCidr) { $script:txtWitnessVsanCidr.Text = $Cfg.WitnessVsanCidr }
  if ($null -ne $Cfg.WitnessShared) { $script:chkWitnessShared.IsChecked = [bool]$Cfg.WitnessShared }
  Write-Log 'Config loaded into UI (passwords not loaded).'
}

function Save-UiConfig { param([Parameter(Mandatory)][string]$Path) (Get-UiConfig | ConvertTo-Json -Depth 8) | Set-Content -Path $Path -Encoding UTF8; Write-Log "Saved config: $Path" }
function Load-UiConfig { param([Parameter(Mandatory)][string]$Path) $cfg = (Get-Content -Raw -Path $Path | ConvertFrom-Json); Apply-UiConfig -Cfg $cfg }
# --- Events ---
$script:window.Add_ContentRendered({
  try {
    if (-not $script:RunDir) { $null = New-RunDir -Base $script:ReportsBase }
    if ($script:txtReports) { $script:txtReports.Text = $script:ReportsBase }
    Write-Log "==== VCF Stretch UI started (v$Global:VCFStretchVersion) ===="
    Write-Log "Run folder: $script:RunDir"
    # uiTimer removed
    Prereq-Check | Out-Null

    # Defaults: ESA + 2 vDS
    $script:cmbStorageType.SelectedIndex = 0
    $script:cmbVdsConfig.SelectedIndex = 1
    Refresh-TopologyUI
  Ensure-DefaultVmnicMapping

Ensure-DefaultNetworkMappings -Force
Set-MappingFieldsReadOnly
 if ($script:txtSddcUser -and [string]::IsNullOrWhiteSpace(($script:txtSddcUser.Text + '').Trim())) { $script:txtSddcUser.Text = 'administrator@vsphere.local' }
  

# Prefill example formats for complex fields (edit as needed)
 try {
   if ($script:txtVmnicMapping) { $script:txtVmnicMapping.Text = Normalize-Arrow ($script:txtVmnicMapping.Text + '') }
   if ($script:txtVdsToNsxUplinkMap) { $script:txtVdsToNsxUplinkMap.Text = Normalize-Arrow ($script:txtVdsToNsxUplinkMap.Text + '') }
   if ($script:txtActiveUplinks) { $script:txtActiveUplinks.Text = Normalize-Arrow ($script:txtActiveUplinks.Text + '') }
 } catch {}
# Prefill example formats for complex fields (edit as needed)
try {
  if ($script:txtVmnicMapping -and [string]::IsNullOrWhiteSpace(($script:txtVmnicMapping.Text + '').Trim())) { $script:txtVmnicMapping.Text = 'vmnic0->VDS1/uplink1; vmnic1->VDS1/uplink1; vmnic2->VDS1/uplink1; vmnic3->VDS1/uplink1' }
  if ($script:txtVdsToNsxUplinkMap -and [string]::IsNullOrWhiteSpace(($script:txtVdsToNsxUplinkMap.Text + '').Trim())) { $script:txtVdsToNsxUplinkMap.Text = 'uplink1->uplink-1; uplink2->uplink-2' }
  if ($script:txtTepPoolName -and [string]::IsNullOrWhiteSpace(($script:txtTepPoolName.Text + '').Trim())) { $script:txtTepPoolName.Text = 'az2-tep-pool' }
  if ($script:txtTepCidr -and [string]::IsNullOrWhiteSpace(($script:txtTepCidr.Text + '').Trim())) { $script:txtTepCidr.Text = '10.0.0.0/24' }
  if ($script:txtTepGateway -and [string]::IsNullOrWhiteSpace(($script:txtTepGateway.Text + '').Trim())) { $script:txtTepGateway.Text = '10.0.0.1' }
  if ($script:txtTepRangeStart -and [string]::IsNullOrWhiteSpace(($script:txtTepRangeStart.Text + '').Trim())) { $script:txtTepRangeStart.Text = '10.0.0.101' }
  if ($script:txtTepRangeEnd -and [string]::IsNullOrWhiteSpace(($script:txtTepRangeEnd.Text + '').Trim())) { $script:txtTepRangeEnd.Text = '10.0.0.132' }
  if ($script:txtUplinkProfileName -and [string]::IsNullOrWhiteSpace(($script:txtUplinkProfileName.Text + '').Trim())) { $script:txtUplinkProfileName.Text = 'az2-uplink-profile' }
  if ($script:txtTransportVlan -and [string]::IsNullOrWhiteSpace(($script:txtTransportVlan.Text + '').Trim())) { $script:txtTransportVlan.Text = '0' }
  if ($script:txtTeamingPolicy -and [string]::IsNullOrWhiteSpace(($script:txtTeamingPolicy.Text + '').Trim())) { $script:txtTeamingPolicy.Text = 'LOADBALANCE_SRCID' }
  if ($script:txtActiveUplinks -and [string]::IsNullOrWhiteSpace(($script:txtActiveUplinks.Text + '').Trim())) { $script:txtActiveUplinks.Text = 'uplink1,uplink2' }
  if ($script:txtStandbyUplinks -and [string]::IsNullOrWhiteSpace(($script:txtStandbyUplinks.Text + '').Trim())) { $script:txtStandbyUplinks.Text = '' }
  if ('' -ne '' -and $script:txtAz2NetworkProfileName -and [string]::IsNullOrWhiteSpace(($script:txtAz2NetworkProfileName.Text + '').Trim())) { $script:txtAz2NetworkProfileName.Text = '' }
  if ('' -ne '' -and $script:txtNsxHostSwitchVds -and [string]::IsNullOrWhiteSpace(($script:txtNsxHostSwitchVds.Text + '').Trim())) { $script:txtNsxHostSwitchVds.Text = '' }
} catch {}
} catch {}
})

if ($script:btnRecheck) {
  $script:btnRecheck.Add_Click({ Prereq-Check | Out-Null })
}
if ($script:btnInstallImpExcel) {
  $script:btnInstallImpExcel.Add_Click({ Ensure-Module -Name 'ImportExcel' | Out-Null; Prereq-Check | Out-Null })
}
if ($script:btnInstallPCLI) {
  $script:btnInstallPCLI.Add_Click({ Ensure-Module -Name 'VMware.PowerCLI' | Out-Null; Prereq-Check | Out-Null })
}


if ($script:btnInstallVCFPCLI) {
 $script:btnInstallVCFPCLI.Add_Click({ Ensure-Module -Name 'VCF.PowerCLI' | Out-Null; Prereq-Check | Out-Null })
}



# Config Save/Load
if ($script:btnBrowseConfig) {
  $script:btnBrowseConfig.Add_Click({
    try {
      $dlg = New-Object System.Windows.Forms.OpenFileDialog
      $dlg.Filter = 'JSON config (*.json)|*.json|All files (*.*)|*.*'
      $dlg.Title = 'Select config file'
      if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:txtConfigPath.Text = $dlg.FileName
      }
    } catch {}
  })
}
if ($script:btnLoadConfig) {
  $script:btnLoadConfig.Add_Click({
    try {
      $p = ($script:txtConfigPath.Text + '').Trim()
      if ([string]::IsNullOrWhiteSpace($p) -or -not (Test-Path $p)) { throw 'Select a config file to load.' }
      Load-UiConfig -Path $p
    } catch {
      Write-Log "Load config failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("Load config failed: $($_.Exception.Message)",'VCF Stretch','OK','Error') | Out-Null
    }
  })
}
if ($script:btnSaveConfig) {
  $script:btnSaveConfig.Add_Click({
    try {
      $dlg = New-Object System.Windows.Forms.SaveFileDialog
      $dlg.Filter = 'JSON config (*.json)|*.json|All files (*.*)|*.*'
      $dlg.Title = 'Save config file'
      $dlg.FileName = 'vcf-stretch-config.json'
      if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:txtConfigPath.Text = $dlg.FileName
        Save-UiConfig -Path $dlg.FileName
      }
    } catch {
      Write-Log "Save config failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("Save config failed: $($_.Exception.Message)",'VCF Stretch','OK','Error') | Out-Null
    }
  })
}
$script:cmbStorageType.Add_SelectionChanged({ Refresh-TopologyUI; Ensure-DefaultNetworkMappings })
$script:cmbVdsConfig.Add_SelectionChanged({ Refresh-TopologyUI; Ensure-DefaultNetworkMappings -Force })
if ($script:chkManualMapOverride) { $script:chkManualMapOverride.Add_Click({ Set-MappingFieldsReadOnly }) }
if ($script:txtNsxHostSwitchVds) { $script:txtNsxHostSwitchVds.Add_LostFocus({ Ensure-DefaultNetworkMappings -Force }) }


if ($script:btnBrowseReports) {
  $script:btnBrowseReports.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select reports/output folder'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $script:ReportsBase = $dlg.SelectedPath
      if ($script:txtReports) { $script:txtReports.Text = $script:ReportsBase }
      $null = New-RunDir -Base $script:ReportsBase
      Write-Log "Reports base set to: $script:ReportsBase"
      Write-Log "New run folder: $script:RunDir"
    }
  })
}

if ($script:btnOpenOut) {
  $script:btnOpenOut.Add_Click({
    try {
      if ($script:RunDir -and (Test-Path $script:RunDir)) { Start-Process $script:RunDir | Out-Null }


if ($script:btnOpenLogFolder) {
 $script:btnOpenLogFolder.Add_Click({
  try {
   if ($Global:LogFile -and (Test-Path $Global:LogFile)) {
    Start-Process explorer.exe ("/select,`"{0}`"" -f $Global:LogFile) | Out-Null
   } elseif ($script:RunDir -and (Test-Path $script:RunDir)) {
    Start-Process $script:RunDir | Out-Null
   }
  } catch {}
 })
}
      elseif ($script:ReportsBase -and (Test-Path $script:ReportsBase)) { Start-Process $script:ReportsBase | Out-Null }
    } catch {}
  })
}

if ($script:btnOpenLog) {
  $script:btnOpenLog.Add_Click({
    try { if ($Global:LogFile -and (Test-Path $Global:LogFile)) { Start-Process notepad.exe $Global:LogFile | Out-Null } } catch {}
  })
}



# vCenter connect/verify (required for Generate/Validate/Execute)
if ($script:btnVerifyVCenter) {
  $script:btnVerifyVCenter.Add_Click({
    Write-Log 'Connect/Verify clicked.'
    try {
      $vc = ($script:txtVCenterFqdn.Text + '').Trim()
      $vu = ($script:txtVCenterUser.Text + '').Trim()
      $vp = ($script:pbVCenterPass.Password + '')
      if ([string]::IsNullOrWhiteSpace($vc) -or [string]::IsNullOrWhiteSpace($vu) -or [string]::IsNullOrWhiteSpace($vp)) {
        throw 'vCenter FQDN, Username, and Password are required.'
      }
      Write-Log "Connecting to vCenter: $vc"
      $script:VCenterServer = Connect-VCenterIfNeeded -Server $vc -User $vu -Pass $vp
      $script:VCenterVerified = $true
      try { if ($script:lblVCenterStatus) { $script:lblVCenterStatus.Text = 'Connected'; $script:lblVCenterStatus.Foreground = [Windows.Media.Brushes]::LightGreen } } catch {}

      try {
        if ($script:LastNetworkQueryResult) { AutoPopulate-FromNetworkQuery -QueryResult $script:LastNetworkQueryResult }
        else { Write-Log 'Tip: Click Collect to pull VCENTER_NSXT_NETWORK_CONFIG and auto-fill some fields.' 'INFO' }
      } catch { Write-Log "Auto-fill from SDDC query failed: $($_.Exception.Message)" 'WARN' }

      try {
        if (Get-Command Get-VDSwitch -ErrorAction SilentlyContinue) {
          $vds = @(Get-VDSwitch | Select-Object -ExpandProperty Name)
          if ($vds -and $vds.Count -gt 0 -and $script:txtNsxHostSwitchVds) {
            if ([string]::IsNullOrWhiteSpace(($script:txtNsxHostSwitchVds.Text + '').Trim())) {
              $script:txtNsxHostSwitchVds.Text = $vds[0]
              Write-Log "Auto-set NSX Host Switch vDS name from vCenter: $($vds[0])"
            }
          }
        }
      } catch { Write-Log "vCenter inventory prefill skipped: $($_.Exception.Message)" 'WARN' }

      # If vmnic mapping is still blank or placeholder, set a safe default now that vDS01 is known.
      Ensure-DefaultVmnicMapping
Ensure-DefaultNetworkMappings -Force
Set-MappingFieldsReadOnly
Write-Log 'vCenter verified.'
      Refresh-TopologyUI
    } catch {
      $script:VCenterVerified = $false
      try { if ($script:lblVCenterStatus) { $script:lblVCenterStatus.Text = 'Connect failed'; $script:lblVCenterStatus.Foreground = [Windows.Media.Brushes]::Tomato } } catch {}
      Write-Log "vCenter verification failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("vCenter verification failed: $($_.Exception.Message)",'VCF Stretch','OK','Error') | Out-Null
      Refresh-TopologyUI
    }
  })
}
# AZ2 host picker (UNASSIGNED/USEABLE)
if ($script:btnPickHosts) {
 $script:btnPickHosts.Add_Click({
  try {
   if (-not $script:SddcSession) { throw 'Connect to SDDC Manager API first.' }
   $picked = Show-HostPicker -Session $script:SddcSession
   if ($picked -and $picked.Count -gt 0) {
    $script:txtAz2Hosts.Text = ($picked -join "`r`n")
    Write-Log ("Picked AZ2 hosts: {0}" -f ($picked -join ', '))
   }
  } catch {
   Write-Log "Host picker failed: $($_.Exception.Message)" 'ERROR'
   [System.Windows.MessageBox]::Show("Host picker failed: $($_.Exception.Message)","VCF Stretch",'OK','Error') | Out-Null
  }
 })
}
if ($script:btnConnect) {
  $script:btnConnect.Add_Click({
    try {
      if (-not (Prereq-Check)) { throw 'Prerequisites not met.' }
      $sddcHost = ($script:txtSddcHost.Text + '').Trim()
      $user = ($script:txtSddcUser.Text + '').Trim()
      $pass = ($script:pbSddcPass.Password + '')
      if ([string]::IsNullOrWhiteSpace($sddcHost) -or [string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
        throw 'SDDC Manager host/username/password required.'
      }

      Write-Log "Connecting to SDDC Manager API: $sddcHost"
      $script:SddcSession = New-SddcToken -SddcHost $sddcHost -Username $user -Password $pass
      Write-Log 'Token acquired.'
 try { if ($script:lblConnStatus) { $script:lblConnStatus.Text = 'Connected'; $script:lblConnStatus.Foreground = [Windows.Media.Brushes]::LightGreen } } catch {}

      # Load clusters into dropdown
      $clusters = Get-Clusters -Session $script:SddcSession
      $script:ClusterMap = @{}
      $script:cmbCluster.Items.Clear()
      foreach($c in $clusters) {
        $name = Coalesce $c.name $c.clusterName
        $id = Coalesce $c.id $c.clusterId
        if (-not $id) { continue }
        $label = if ($name) { "$name ($id)" } else { "$id" }
        $null = $script:cmbCluster.Items.Add($label)
        $script:ClusterMap[$label] = $id
      }
      if ($script:cmbCluster.Items.Count -gt 0) { $script:cmbCluster.SelectedIndex = 0 }

      Refresh-TopologyUI
      Write-Log "Clusters loaded: $($script:cmbCluster.Items.Count)"
 Apply-PendingClusterSelection

    } catch {
      try { if ($script:lblConnStatus) { $script:lblConnStatus.Text = 'Connect failed'; $script:lblConnStatus.Foreground = [Windows.Media.Brushes]::Tomato } } catch {}
Write-Log "Connect failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("Connect failed: $($_.Exception.Message)","VCF Stretch",'OK','Error') | Out-Null
    }
  })
}

if ($script:btnCollect) {
  $script:btnCollect.Add_Click({
    try {
      if (-not $script:SddcSession) { throw 'Connect to SDDC Manager API first.' }
      $cid = Get-SelectedClusterId
      if (-not $cid) { throw 'Select a cluster.' }

      Write-Log "Collecting network configuration via SDDC Manager API query (VCENTER_NSXT_NETWORK_CONFIG)..."
      $qid = Start-ClusterNetworkQuery -Session $script:SddcSession -ClusterId $cid
      if (-not $qid) { throw 'Network query did not return a queryId.' }
      $res = Get-ClusterNetworkQueryResult -Session $script:SddcSession -ClusterId $cid -QueryId $qid
$script:LastNetworkQueryResult = $res
$out = Join-Path $script:RunDir "ClusterNetworkConfig-$cid.json"
($res | ConvertTo-Json -Depth 30) | Set-Content -Path $out -Encoding UTF8
Write-Log "Saved network query output: $out"
try { AutoPopulate-FromNetworkQuery -QueryResult $res } catch { Write-Log "Auto-fill failed: $($_.Exception.Message)" 'WARN' }
      # Force deterministic network mappings after Collect (Collect may not populate mapping fields)
Ensure-DefaultNetworkMappings -Force
Set-MappingFieldsReadOnly
# Note: We do not attempt to auto-parse into fields in v0.9 (varies by environment).
      # Operators can use the saved JSON to assist, and validations will catch mismatches.

      [System.Windows.MessageBox]::Show("Network config collected and saved to run folder:\n$out","VCF Stretch",'OK','Information') | Out-Null

    } catch {
      Write-Log "Collect failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("Collect failed: $($_.Exception.Message)","VCF Stretch",'OK','Error') | Out-Null
    }
  })
}

function Build-SpecFromUi {
  Ensure-DefaultNetworkMappings -Force
  if (-not $script:SddcSession) { throw 'Connect to SDDC Manager API first.' }
 if (-not $script:VCenterVerified) { throw 'vCenter connection/verification is required. Connect to vCenter and click Connect/Verify.' }

  $top = Get-SelectedTopology
  if (-not $top.IsSupported) { throw 'Selected topology not supported by template.' }

  $clusterId = Get-SelectedClusterId
  if (-not $clusterId) { throw 'Select a cluster.' }

  $clusterLabel = $script:cmbCluster.SelectedItem.ToString()
  $clusterName = if ($clusterLabel -match '^(.+?)\s*\(') { $Matches[1].Trim() } else { $clusterLabel }

  $az1 = ($script:txtAz1.Text + '').Trim()
  $az2 = ($script:txtAz2.Text + '').Trim()
  if ([string]::IsNullOrWhiteSpace($az1) -or [string]::IsNullOrWhiteSpace($az2)) { throw 'AZ1 and AZ2 names are required.' }

  $hosts = Get-Az2HostsFromUi
  if ($hosts.Count -lt 1) { throw 'At least one AZ2 host FQDN is required.' }

  $az2np = ($script:txtAz2NetworkProfileName.Text + '').Trim()
 $suffix = ((($script:txtNameSuffix.Text + '')).Trim()); if ([string]::IsNullOrWhiteSpace($suffix)) { $suffix = '01' }
 if ([string]::IsNullOrWhiteSpace($az2np)) { $az2np = ("$az2-network-profile-" + $suffix); $script:txtAz2NetworkProfileName.Text = $az2np }
  
  $vmnicMap = (Normalize-Arrow ($script:txtVmnicMapping.Text + '')).Trim()
  if ([string]::IsNullOrWhiteSpace($vmnicMap)) { throw 'Standard vmnic mapping is required.' }

  $nsxVds = ($script:txtNsxHostSwitchVds.Text + '').Trim()
  if ([string]::IsNullOrWhiteSpace($nsxVds)) { throw 'NSX Host Switch vDS name is required.' }

  $tepPool = ($script:txtTepPoolName.Text + '').Trim()
 if ([string]::IsNullOrWhiteSpace($tepPool)) { $tepPool = ("$az2np-ip-pool-host-" + $suffix); $script:txtTepPoolName.Text = $tepPool }
 try { if ($script:lblDerivedIpPoolName) { $script:lblDerivedIpPoolName.Text = $tepPool } } catch {}
  $tepCidr = ($script:txtTepCidr.Text + '').Trim()
  $tepGw = ($script:txtTepGateway.Text + '').Trim()
  $tepStart = ($script:txtTepRangeStart.Text + '').Trim()
  $tepEnd = ($script:txtTepRangeEnd.Text + '').Trim()
  if (@($tepPool,$tepCidr,$tepGw,$tepStart,$tepEnd) | Where-Object { [string]::IsNullOrWhiteSpace($_) }) {
    throw 'TEP IP pool fields (name/cidr/gateway/start/end) are required.'
  }

  $uplinkProf = ($script:txtUplinkProfileName.Text + '').Trim()
 if ([string]::IsNullOrWhiteSpace($uplinkProf)) { $uplinkProf = ("$az2np-uplink-profile-" + $suffix); $script:txtUplinkProfileName.Text = $uplinkProf }
 try { if ($script:lblDerivedUplinkProfileName) { $script:lblDerivedUplinkProfileName.Text = $uplinkProf } } catch {}
  $transportVlan = ($script:txtTransportVlan.Text + '').Trim()
  $teamPolicy = ($script:txtTeamingPolicy.Text + '').Trim()
  $activeU = ($script:txtActiveUplinks.Text + '').Trim()
  $standbyU = ($script:txtStandbyUplinks.Text + '').Trim()



# --- vDS uplink -> NSX uplink mapping (must be present) ---
# Reconcile uplink names: Active/Standby must reference NSX uplink names (nsxUplinkName)
$pairs = Parse-UplinkMapping $mapText
$nsxNames = @($pairs | Select-Object -ExpandProperty nsxUplinkName)

$activeU  = ($script:txtActiveUplinks.Text + '').Trim()
$standbyU = ($script:txtStandbyUplinks.Text + '').Trim()

$activeList  = @(Parse-CommaList (Normalize-Arrow $activeU))
$standbyList = @(Parse-CommaList (Normalize-Arrow $standbyU))

$activeConverted = foreach ($a in $activeList) {
  $m = $pairs | Where-Object { $_.vdsUplinkName -eq $a } | Select-Object -First 1
  if ($m) { $m.nsxUplinkName } else { $a }
}

$standbyConverted = foreach ($s in $standbyList) {
  $m = $pairs | Where-Object { $_.vdsUplinkName -eq $s } | Select-Object -First 1
  if ($m) { $m.nsxUplinkName } else { $s }
}

if ($activeConverted.Count -gt 0) {
  $activeU = ($activeConverted -join ',')
  $script:txtActiveUplinks.Text = $activeU
  Write-Log "Normalized Active Uplinks to NSX names: $activeU"
}

if ($standbyConverted.Count -gt 0) {
  $standbyU = ($standbyConverted -join ',')
  $script:txtStandbyUplinks.Text = $standbyU
  Write-Log "Normalized Standby Uplinks to NSX names: $standbyU"
}

# If ActiveUplinks is empty, default to NSX uplinks from mapping
if ([string]::IsNullOrWhiteSpace($activeU) -and $nsxNames.Count -gt 0) {
  $activeU = ($nsxNames -join ',')
  $script:txtActiveUplinks.Text = $activeU
  Write-Log "Defaulted Active Uplinks from mapping: $activeU"
}

if (@($uplinkProf,$transportVlan,$teamPolicy,$activeU) | Where-Object { [string]::IsNullOrWhiteSpace($_) }) {
    throw 'Uplink profile fields (name/transport VLAN/teaming policy/active uplinks) are required.'
  }

  $mapText = (Normalize-Arrow ($script:txtVdsToNsxUplinkMap.Text + '')).Trim()
  if ([string]::IsNullOrWhiteSpace($mapText)) { throw 'vDS uplink to NSX uplink mapping is required.' }

  $wFqdn = ($script:txtWitnessFqdn.Text + '').Trim()
  $wIp = ($script:txtWitnessVsanIp.Text + '').Trim()
  $wCidr = ($script:txtWitnessVsanCidr.Text + '').Trim()
  if (@($wFqdn,$wIp,$wCidr) | Where-Object { [string]::IsNullOrWhiteSpace($_) }) {
    throw 'Witness fields (FQDN, vSAN IP, vSAN CIDR) are required.'
  }

  $spec = New-ClusterStretchSpec -Session $script:SddcSession -ClusterId $clusterId -StorageType $top.StorageType -VdsCount $top.VdsCount `
    -Az1Name $az1 -Az2Name $az2 -Az2HostsFqdn $hosts -Az2NetworkProfileName $az2np -Az2NetworkProfileIsDefault ([bool]$script:chkAz2NetworkProfileDefault.IsChecked) `
    -VmnicMappingText $vmnicMap -NsxHostSwitchVdsName $nsxVds -TepPoolName $tepPool -TepCidr $tepCidr -TepGateway $tepGw -TepRangeStart $tepStart -TepRangeEnd $tepEnd `
    -UplinkProfileName $uplinkProf -TransportVlan ([int]$transportVlan) -TeamingPolicy $teamPolicy -ActiveUplinksCsv $activeU -StandbyUplinksCsv $standbyU `
    -VdsToNsxUplinkMapText $mapText -DeployWithoutLicenseKeys ([bool]$script:chkDeployNoLic.IsChecked) -IsEdgeClusterConfiguredForMultiAZ ([bool]$script:chkEdgeMultiAZ.IsChecked) `
    -WitnessFqdn $wFqdn -WitnessVsanIp $wIp -WitnessVsanCidr $wCidr -WitnessTrafficSharedWithVsanTraffic ([bool]$script:chkWitnessShared.IsChecked)

  return [pscustomobject]@{ Spec=$spec; ClusterName=$clusterName; ClusterId=$clusterId; Hosts=$hosts; Az1=$az1; Az2=$az2; SheetName=$top.SheetName }
}

if ($script:btnGenerate) {
  $script:btnGenerate.Add_Click({
    try {
      if (-not $script:RunDir) { $null = New-RunDir -Base $script:ReportsBase }
      $ctx = Build-SpecFromUi
      $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
      $outJson = Join-Path $script:RunDir ("clusterStretchSpec_$($ctx.ClusterId)_$ts.json")
      $outWrap = Join-Path $script:RunDir ("clusterUpdateSpec_validationWrapper_$($ctx.ClusterId)_$ts.json")

      ($ctx.Spec | ConvertTo-Json -Depth 30) | Set-Content -Path $outJson -Encoding UTF8
      (@{ clusterUpdateSpec = $ctx.Spec } | ConvertTo-Json -Depth 30) | Set-Content -Path $outWrap -Encoding UTF8

      Write-Log "Wrote JSON: $outJson"
      Write-Log "Wrote validation wrapper: $outWrap"
      [System.Windows.MessageBox]::Show("Generated JSON outputs in run folder:`n$outJson`n$outWrap","VCF Stretch",'OK','Information') | Out-Null
    } catch {
      Write-Log "Generate failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("Generate failed: $($_.Exception.Message)","VCF Stretch",'OK','Error') | Out-Null
    }
  })
}
if ($script:btnValidate) {
  $script:btnValidate.Add_Click({
    try {
      $ctx = Build-SpecFromUi
      $wrap = @{ clusterUpdateSpec = $ctx.Spec }
      Write-Log "Validating stretch spec via POST /v1/clusters/$($ctx.ClusterId)/validations..."
      $res = Invoke-SddcApi -Session $script:SddcSession -Method POST -Path ("/v1/clusters/{0}/validations" -f $ctx.ClusterId) -Body $wrap
      $out = Join-Path $script:RunDir ("ValidationResponse_$($ctx.ClusterId)_" + (Get-Date -Format 'yyyyMMdd-HHmmss') + ".json")
      ($res | ConvertTo-Json -Depth 30) | Set-Content -Path $out -Encoding UTF8
      Write-Log "Validation response saved: $out"
      [System.Windows.MessageBox]::Show("Validation request submitted. Response saved to:\n$out","VCF Stretch",'OK','Information') | Out-Null
    } catch {
      Write-Log "Validate failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("Validate failed: $($_.Exception.Message)","VCF Stretch",'OK','Error') | Out-Null
    }
  })
}

if ($script:btnExecute) {
  $script:btnExecute.Add_Click({
    try {
      $ctx = Build-SpecFromUi
      $msg = "This will initiate a cluster stretch via PATCH /v1/clusters/$($ctx.ClusterId).\nContinue?"
      $ans = [System.Windows.MessageBox]::Show($msg,'VCF Stretch','YesNo','Warning')
      if ($ans -ne 'Yes') { return }

      Write-Log "Executing stretch via PATCH /v1/clusters/$($ctx.ClusterId)..."
      $res = Invoke-SddcApi -Session $script:SddcSession -Method PATCH -Path ("/v1/clusters/{0}" -f $ctx.ClusterId) -Body $ctx.Spec
      $out = Join-Path $script:RunDir ("ExecuteResponse_$($ctx.ClusterId)_" + (Get-Date -Format 'yyyyMMdd-HHmmss') + ".json")
      ($res | ConvertTo-Json -Depth 30) | Set-Content -Path $out -Encoding UTF8
      Write-Log "Execute response saved: $out"
      [System.Windows.MessageBox]::Show("Execute request submitted. Response saved to:\n$out","VCF Stretch",'OK','Information') | Out-Null
    } catch {
      Write-Log "Execute failed: $($_.Exception.Message)" 'ERROR'
      [System.Windows.MessageBox]::Show("Execute failed: $($_.Exception.Message)","VCF Stretch",'OK','Error') | Out-Null
    }
  })
}

if ($script:btnClose) {
  $script:btnClose.Add_Click({
    try { $script:window.Close() } catch {}
  })
}

# Final: show UI
# --- Overrides: robust Save/Load (exclude passwords), normalize arrows, null-safe assignments ---
function Apply-UiConfig {
  param([Parameter(Mandatory)]$Cfg)
  try {
    Write-Log 'Applying loaded config to UI (passwords excluded).'

    if ($null -ne $Cfg.VCenterFqdn) { $script:txtVCenterFqdn.Text = $Cfg.VCenterFqdn }
    if ($null -ne $Cfg.VCenterUser) { $script:txtVCenterUser.Text = $Cfg.VCenterUser }
    if ($null -ne $Cfg.SddcHost) { $script:txtSddcHost.Text = $Cfg.SddcHost }
    if ($null -ne $Cfg.SddcUser) { $script:txtSddcUser.Text = $Cfg.SddcUser }

    if ($null -ne $Cfg.Az1) { $script:txtAz1.Text = $Cfg.Az1 }
    if ($null -ne $Cfg.Az2) { $script:txtAz2.Text = $Cfg.Az2 }
    if ($null -ne $Cfg.Az2Hosts) { $script:txtAz2Hosts.Text = $Cfg.Az2Hosts }

    if ($null -ne $Cfg.DeployNoLic) { $script:chkDeployNoLic.IsChecked = [bool]$Cfg.DeployNoLic }
    if ($null -ne $Cfg.EdgeMultiAz) { $script:chkEdgeMultiAZ.IsChecked = [bool]$Cfg.EdgeMultiAz }

    if ($null -ne $Cfg.Az2NetworkProfileName) { $script:txtAz2NetworkProfileName.Text = $Cfg.Az2NetworkProfileName }
    if ($null -ne $Cfg.Az2NetworkProfileDefault) { $script:chkAz2NetworkProfileDefault.IsChecked = [bool]$Cfg.Az2NetworkProfileDefault }
    if ($null -ne $Cfg.NameSuffix -and $script:txtNameSuffix) { $script:txtNameSuffix.Text = $Cfg.NameSuffix }

    if ($null -ne $Cfg.VmnicMapping) { $script:txtVmnicMapping.Text = Normalize-Arrow $Cfg.VmnicMapping }
    if ($null -ne $Cfg.NsxHostSwitchVds) { $script:txtNsxHostSwitchVds.Text = $Cfg.NsxHostSwitchVds }
    if ($null -ne $Cfg.VdsToNsxUplinkMap) { $script:txtVdsToNsxUplinkMap.Text = Normalize-Arrow $Cfg.VdsToNsxUplinkMap }

    if ($null -ne $Cfg.TepCidr) { $script:txtTepCidr.Text = $Cfg.TepCidr }
    if ($null -ne $Cfg.TepGateway) { $script:txtTepGateway.Text = $Cfg.TepGateway }
    if ($null -ne $Cfg.TepRangeStart) { $script:txtTepRangeStart.Text = $Cfg.TepRangeStart }
    if ($null -ne $Cfg.TepRangeEnd) { $script:txtTepRangeEnd.Text = $Cfg.TepRangeEnd }

    if ($null -ne $Cfg.TransportVlan) { $script:txtTransportVlan.Text = $Cfg.TransportVlan }
    if ($null -ne $Cfg.TeamingPolicy) { $script:txtTeamingPolicy.Text = $Cfg.TeamingPolicy }
    if ($null -ne $Cfg.ActiveUplinks) { $script:txtActiveUplinks.Text = Normalize-Arrow $Cfg.ActiveUplinks }
    if ($null -ne $Cfg.StandbyUplinks) { $script:txtStandbyUplinks.Text = Normalize-Arrow $Cfg.StandbyUplinks }

    if ($null -ne $Cfg.WitnessFqdn) { $script:txtWitnessFqdn.Text = $Cfg.WitnessFqdn }
    if ($null -ne $Cfg.WitnessVsanIp) { $script:txtWitnessVsanIp.Text = $Cfg.WitnessVsanIp }
    if ($null -ne $Cfg.WitnessVsanCidr) { $script:txtWitnessVsanCidr.Text = $Cfg.WitnessVsanCidr }
    if ($null -ne $Cfg.WitnessShared) { $script:chkWitnessShared.IsChecked = [bool]$Cfg.WitnessShared }

    try { if ($null -ne $Cfg.ClusterLabel) { $script:PendingClusterLabel = ($Cfg.ClusterLabel + '').Trim() } } catch {}

    # If using 2 vDS and vmnic mapping is blank/placeholder, seed it.
    Ensure-DefaultVmnicMapping

    Write-Log 'Config applied.'
  } catch {
    Write-Log "Apply config failed: $($_.Exception.Message)" 'ERROR'
  }
}

function Apply-PendingClusterSelection {
  try {
    if ($script:PendingClusterLabel -and $script:cmbCluster -and $script:cmbCluster.Items.Count -gt 0) {
      for ($i=0; $i -lt $script:cmbCluster.Items.Count; $i++) {
        if (($script:cmbCluster.Items[$i].ToString()) -eq $script:PendingClusterLabel) {
          $script:cmbCluster.SelectedIndex = $i
          $script:PendingClusterLabel = $null
          Write-Log 'Applied saved cluster selection.'
          break
        }
      }
    }
  } catch {}
}

function Load-UiConfig {
  param([Parameter(Mandatory)][string]$Path)
  Write-Log "Loading config: $Path"
  $raw = Get-Content -Raw -Path $Path
  $cfg = $raw | ConvertFrom-Json
  Apply-UiConfig -Cfg $cfg
  Apply-PendingClusterSelection
}

function Save-UiConfig {
  param([Parameter(Mandatory)][string]$Path)
  $cfg = Get-UiConfig
  # normalize arrow encodings before writing
  try {
    if ($cfg.VmnicMapping) { $cfg.VmnicMapping = Normalize-Arrow $cfg.VmnicMapping }
    if ($cfg.VdsToNsxUplinkMap) { $cfg.VdsToNsxUplinkMap = Normalize-Arrow $cfg.VdsToNsxUplinkMap }
    if ($cfg.ActiveUplinks) { $cfg.ActiveUplinks = Normalize-Arrow $cfg.ActiveUplinks }
    if ($cfg.StandbyUplinks) { $cfg.StandbyUplinks = Normalize-Arrow $cfg.StandbyUplinks }
  } catch {}
  ($cfg | ConvertTo-Json -Depth 8) | Set-Content -Path $Path -Encoding UTF8
  Write-Log "Saved config: $Path"
}
# Final: show UI

# Final: show UI
$null = $script:window.ShowDialog() # fixed
# SIG # Begin signature block
# MIIHgQYJKoZIhvcNAQcCoIIHcjCCB24CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBMUAm8HnWbxTDO
# iKu/8uoavNSdEF130ZZnLuPjG3TdNqCCBFIwggROMIICtqADAgECAhAUjgWPDHlm
# m0bNsrQuJnTGMA0GCSqGSIb3DQEBCwUAMD8xPTA7BgNVBAMMNFZDRlN0cmV0Y2gg
# TG9jYWwgQ29kZSBTaWduaW5nICh4YWRtaW5ASE9NRU9GRklDRUxBQikwHhcNMjYw
# MjIyMTY0MTM5WhcNMzEwMjIyMTY1MTM5WjA/MT0wOwYDVQQDDDRWQ0ZTdHJldGNo
# IExvY2FsIENvZGUgU2lnbmluZyAoeGFkbWluQEhPTUVPRkZJQ0VMQUIpMIIBojAN
# BgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEAxpNbdm/+gOiX5RxZ95i3S0Xs5vup
# Q0p7+0ImW6L+beFB7BI0bNTlyHrOY/c0HGMh9648xogcoOxInMzlx+0SH9nscyMh
# e3F0ROpOSW6emSM9tHwSIc8g743VaVq0PtRFcbr5iRovsvVysTv7C4RvnMoCSulB
# UmmQM2wxk6jFw5vcgcFfuGvwnOCcSPZuEZBDrms6gdZK84eC3jwlN717b3DzCpx7
# JKmrnW2B6L4nXNHHDyRrMPZ/4OErmnTnAZieUIbFYrOGLDBiqDANW7uxgCSq68y9
# gw/B75yuuG6J42TRR3fMp87F3TR33NlWIuRChYQ2xlYECsg8sBKSBGB2p7nrlL8P
# j8e+tfTrbLILQbYxHzCI+OJWVVfj9fFmkX4z8SSqgvUxrf0K9BW82HZqwUxvDO+n
# hqooPxSHJtD/IKGfFU9lfJqn1OhjDTOr4ar7d7kUgnbYSc0cKtrGzuWvGDPBRPWy
# o4p6ALerREJcRysDtP9lD3EyUCE0fj0NVYiJAgMBAAGjRjBEMA4GA1UdDwEB/wQE
# AwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUUTRK54Yiw5w5lEN9
# vQpo+IQdjRUwDQYJKoZIhvcNAQELBQADggGBAE8kZhpjVlNI5UnMG6D3pe/SYmDm
# jGVmKhNftncJJkAjoUfbF2i4QXjZ9bkOccs5OMjVToUInw0LSxTdv/Vgck90OrSM
# OOpZpPCKtKD306zXmz9Pmn6AHYal0BM64YA5UEHX1t6Hdp4FggEnYwwrvECnGSuF
# 0JFHCE1IiHAK60/eOzwktZN3Q65H3+ypsUnkUKt0N8Bh4/qWWTMAGFzGFCggSoqu
# z9GPSaAU+QxIJvEhgruUwdVuXiYrps83EWOnIf6rgVgjdw9NOCRk4T67dF7Jl73a
# o95zMDzvyoDbmMc/reWa+FZ3kv0AkTWUMEv7i2JL1umoN4qAaF24UYkvSztSc/NZ
# 2uDlaYudhkSz1XFw48u9ho1lBaRhiope2tXiys9/LBtv2bBUll6Z74Xbvh81KubI
# n+s2Nrtq2T3Rw3B0cDCZ1VJ7Vk+78sXQRvfOEJ+oFFCiLPQKDTczhSYA9f/BDslx
# RwapmiqECo6BG7zgTKovft2jkQSSEIB31jIBkzGCAoUwggKBAgEBMFMwPzE9MDsG
# A1UEAww0VkNGU3RyZXRjaCBMb2NhbCBDb2RlIFNpZ25pbmcgKHhhZG1pbkBIT01F
# T0ZGSUNFTEFCKQIQFI4Fjwx5ZptGzbK0LiZ0xjANBglghkgBZQMEAgEFAKCBhDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEi
# BCBNqfb3nAAchp+0kxeBMwvwnXELtQmvB9h01AMv8B+MLzANBgkqhkiG9w0BAQEF
# AASCAYCY1fIIN3Y50otoKi/XNqr6fI1hPZ2Ct4aqqC2N7BiP2T9RKI5lhMJqlS8z
# z3g7/vh5s94ceC3a81NwigYV58qui+1So+3PqaZjsr6Hep5ZwkJQSlZgjCy7kQt5
# 2gnSgUXoHTqzHhpdAt15RhDO6PjZ3gHupS2GQpjoIryOXrWrQNMl19dU71n0YZ6k
# R4EW6aFW+uL5JY7MBx+ZkjIoj49/fHZrYF5iJ/TpCZQiJ+UTa+nCFrqKizfrjvMF
# rjOiyeLW18vCOGQGCKcbkjNLTvWELovZ6SLyKR6lz/7EJSsvKu1kpfRG+f0tDpkn
# NVBfAkjOVsQvDHc3cxSEnvNVxodwdS03yvMehMrhFmpAmnASeSyVNrtXpuIzr4bX
# qozj0frDgJYkGyju9+8gDuFe9KXHuXsbSXAO257fVxVxz5LLJfgjDaOTPEFDOC0R
# D+2u8YRgva+UBx6+GM3zmrOEdnKo/L3qSj65w01FpSFI22EFJHefR8cbZYct9+Kr
# DfHgsg4=
# SIG # End signature block
