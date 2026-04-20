#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$NonInteractive,
    [string]$ReportName
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Continue'

$script:NmBanner = @(
    ' _   _      _                      _   __  __       _        _      ',
    '| \ | | ___| |___      _____  _ __| | |  \/  | __ _| |_ _ __(_)_  __',
    "|  \| |/ _ \ __\ \ /\ / / _ \| '__| | | |\/| |/ _` | __| '__| \ \/ /",
    '| |\  |  __/ |_ \ V  V / (_) | |  | | | |  | | (_| | |_| |  | |>  < ',
    '|_| \_|\___|\__| \_/\_/ \___/|_|  |_| |_|  |_|\__,_|\__|_|  |_/_/\_\',
    '  Camera and recorder deployment survey for unknown sites'
)

$script:NmKnownPorts = @(80, 443, 554, 8000, 8080, 8899, 37777, 5000, 7001)
$script:NmState = [ordered]@{
    SessionName     = $null
    ReportRoot      = $null
    SelectedAdapter = $null
    SubnetInfo      = $null
    LocalSnapshot   = $null
    Inventory       = @()
    LastBrief       = ''
    LastSurveyAt    = $null
    LastPortScanAt  = $null
}

function Get-NmTimestamp {
    Get-Date -Format 'yyyyMMdd_HHmmss'
}

function Test-NmCommandAvailable {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Test-NmIsAdmin {
    try {
        $current = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($current)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Write-NmStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warn','Error','Success')][string]$Level = 'Info'
    )

    $color = switch ($Level) {
        'Warn' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
        default { 'Gray' }
    }

    Write-Host ('[{0}] {1}' -f $Level.ToUpper(), $Message) -ForegroundColor $color
}

function Write-NmSection {
    param([Parameter(Mandatory)][string]$Title)
    Write-Host ''
    Write-Host ('--- {0} ---' -f $Title) -ForegroundColor Cyan
}

function Pause-Nm {
    Read-Host 'Press Enter to continue'
}

function ConvertTo-NmSafeFolderName {
    param([Parameter(Mandatory)][string]$Name)
    $safe = $Name.Trim()
    $safe = $safe -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace '\s+', '_'
    return $safe.Trim('.')
}

function Get-NmReportBaseRoot {
    $portableRoot = Join-Path -Path $PSScriptRoot -ChildPath 'NetworkMatrix_Reports'
    try {
        if (-not (Test-Path -Path $portableRoot)) {
            New-Item -Path $portableRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
            return $portableRoot
        }

        $probe = Join-Path -Path $portableRoot -ChildPath ("probe_{0}.tmp" -f ([guid]::NewGuid().ToString('N')))
        Set-Content -Path $probe -Value 'ok' -Encoding ASCII -ErrorAction Stop
        Remove-Item -Path $probe -Force -ErrorAction SilentlyContinue
        return $portableRoot
    }
    catch {
        return (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'NetworkMatrix\Reports')
    }
}

function Initialize-NmReportSession {
    param(
        [AllowNull()][string]$Name,
        [switch]$NonInteractive
    )

    $baseRoot = Get-NmReportBaseRoot
    if (-not (Test-Path -Path $baseRoot)) {
        New-Item -Path $baseRoot -ItemType Directory -Force | Out-Null
    }

    $rawName = if ([string]::IsNullOrWhiteSpace($Name)) {
        'Survey_{0}' -f (Get-NmTimestamp)
    }
    else {
        $Name
    }

    $safeName = ConvertTo-NmSafeFolderName -Name $rawName
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = 'Survey_{0}' -f (Get-NmTimestamp)
    }

    $target = Join-Path -Path $baseRoot -ChildPath $safeName
    try {
        if (-not (Test-Path -Path $target)) {
            New-Item -Path $target -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        $script:NmState.SessionName = $safeName
        $script:NmState.ReportRoot = $target
        if (-not $NonInteractive) {
            Write-NmStatus -Level Success -Message "Reports will save to: $target"
        }
    }
    catch {
        throw "Unable to create report session '$safeName': $($_.Exception.Message)"
    }
}

function Save-NmTextReport {
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Text,
        [string]$Extension = 'txt'
    )

    if (-not $script:NmState.ReportRoot) {
        Initialize-NmReportSession -Name $ReportName -NonInteractive
    }

    $safePrefix = ($Prefix -replace '[^a-zA-Z0-9_-]', '_')
    $path = Join-Path -Path $script:NmState.ReportRoot -ChildPath ('{0}_{1}.{2}' -f $safePrefix, (Get-NmTimestamp), $Extension)
    Set-Content -Path $path -Value $Text -Encoding UTF8
    return $path
}

function Save-NmJsonReport {
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)]$Object
    )

    $json = $Object | ConvertTo-Json -Depth 8
    Save-NmTextReport -Prefix $Prefix -Text $json -Extension 'json'
}

function Save-NmCsvReport {
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)]$Object
    )

    if (-not $script:NmState.ReportRoot) {
        Initialize-NmReportSession -Name $ReportName -NonInteractive
    }

    $safePrefix = ($Prefix -replace '[^a-zA-Z0-9_-]', '_')
    $path = Join-Path -Path $script:NmState.ReportRoot -ChildPath ('{0}_{1}.csv' -f $safePrefix, (Get-NmTimestamp))
    $Object | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    return $path
}

function Show-NmHeader {
    Clear-Host
    foreach ($line in $script:NmBanner) {
        Write-Host $line -ForegroundColor Magenta
    }
    Write-Host ('=' * 78) -ForegroundColor DarkGray
    Write-Host (' Session: {0}  |  Admin: {1}' -f $(if ($script:NmState.SessionName) { $script:NmState.SessionName } else { 'not started' }), $(if (Test-NmIsAdmin) { 'Yes' } else { 'No' })) -ForegroundColor Gray
    Write-Host (' Reports: {0}' -f $(if ($script:NmState.ReportRoot) { $script:NmState.ReportRoot } else { '(not initialized)' })) -ForegroundColor Gray
    if ($script:NmState.SelectedAdapter) {
        Write-Host (' Adapter: {0}  |  Local IP: {1}/{2}  |  Gateway: {3}' -f $script:NmState.SelectedAdapter.InterfaceAlias, $script:NmState.SelectedAdapter.IPv4Address, $script:NmState.SelectedAdapter.PrefixLength, $(if ($script:NmState.SelectedAdapter.DefaultGateway) { $script:NmState.SelectedAdapter.DefaultGateway } else { 'none' })) -ForegroundColor Gray
    }
    else {
        Write-Host ' Adapter: (not selected)' -ForegroundColor Gray
    }
    Write-Host (' Inventory: {0} hosts  |  Last Survey: {1}' -f @($script:NmState.Inventory).Count, $(if ($script:NmState.LastSurveyAt) { $script:NmState.LastSurveyAt.ToString('yyyy-MM-dd HH:mm:ss') } else { 'never' })) -ForegroundColor Gray
    Write-Host ('=' * 78) -ForegroundColor DarkGray
}

function ConvertTo-NmUInt32IP {
    param([Parameter(Mandatory)][string]$IPAddress)
    $bytes = ([System.Net.IPAddress]::Parse($IPAddress)).GetAddressBytes()
    [Array]::Reverse($bytes)
    return [BitConverter]::ToUInt32($bytes, 0)
}

function ConvertFrom-NmUInt32IP {
    param([Parameter(Mandatory)][uint32]$Value)
    $bytes = [BitConverter]::GetBytes($Value)
    [Array]::Reverse($bytes)
    return ([System.Net.IPAddress]::new($bytes)).ToString()
}

function Get-NmMaskFromPrefix {
    param([Parameter(Mandatory)][int]$PrefixLength)
    if ($PrefixLength -lt 0 -or $PrefixLength -gt 32) { throw 'Invalid prefix length.' }
    if ($PrefixLength -eq 0) { return '0.0.0.0' }
    $mask = [uint32]::MaxValue -shl (32 - $PrefixLength)
    return (ConvertFrom-NmUInt32IP -Value $mask)
}

function Get-NmSubnetInfo {
    param(
        [Parameter(Mandatory)][string]$IPAddress,
        [Parameter(Mandatory)][int]$PrefixLength
    )

    $ipInt = ConvertTo-NmUInt32IP -IPAddress $IPAddress
    $maskInt = if ($PrefixLength -eq 0) { [uint32]0 } else { [uint32]::MaxValue -shl (32 - $PrefixLength) }
    $networkInt = $ipInt -band $maskInt
    $broadcastInt = $networkInt + ([uint32]([math]::Pow(2, (32 - $PrefixLength)) - 1))
    $hostCount = if ($PrefixLength -ge 31) { 0 } else { [int64]([math]::Pow(2, (32 - $PrefixLength)) - 2) }

    [pscustomobject]@{
        IPAddress        = $IPAddress
        PrefixLength     = $PrefixLength
        SubnetMask       = Get-NmMaskFromPrefix -PrefixLength $PrefixLength
        NetworkAddress   = ConvertFrom-NmUInt32IP -Value $networkInt
        BroadcastAddress = ConvertFrom-NmUInt32IP -Value $broadcastInt
        FirstUsable      = if ($hostCount -gt 0) { ConvertFrom-NmUInt32IP -Value ($networkInt + 1) } else { $null }
        LastUsable       = if ($hostCount -gt 0) { ConvertFrom-NmUInt32IP -Value ($broadcastInt - 1) } else { $null }
        HostCount        = $hostCount
        NetworkInt       = $networkInt
        BroadcastInt     = $broadcastInt
    }
}

function Resolve-NmMacVendor {
    param([AllowNull()][string]$MacAddress)
    if ([string]::IsNullOrWhiteSpace($MacAddress)) { return '' }

    $prefix = ($MacAddress -replace '[^0-9A-Fa-f]', '').ToUpper()
    if ($prefix.Length -lt 6) { return '' }
    $prefix = $prefix.Substring(0, 6)

    $map = @{
        'B0C5CA' = 'Hikvision'
        'BCAD28' = 'Hikvision'
        'EC172F' = 'Hikvision'
        '9002A9' = 'Dahua'
        '3C2A2F' = 'Dahua'
        'D4E0B0' = 'Axis'
        'ACCC8E' = 'Axis'
        '00085D' = 'Axis'
        '000F7C' = 'Hanwha'
        'B4A36B' = 'Hanwha'
        'D89EF3' = 'Ubiquiti'
        '7483C2' = 'Ubiquiti'
        'FCECDA' = 'Ubiquiti'
        'CC2D21' = 'TP-Link'
        'F4F26D' = 'Aruba'
        '3C5282' = 'Cisco'
        '001560' = 'Cisco'
        '000C29' = 'VMware'
        '000569' = 'VMware'
        '080027' = 'VirtualBox'
        '00155D' = 'Hyper-V'
        'E4956E' = 'Reolink'
        'FCD733' = 'Uniview'
    }

    if ($map.ContainsKey($prefix)) { return $map[$prefix] }
    return ''
}

function Get-NmAdapterCandidates {
    $items = New-Object System.Collections.Generic.List[object]

    if (Test-NmCommandAvailable -Name Get-NetIPConfiguration) {
        $cfg = @()
        try {
            $cfg = @(Get-NetIPConfiguration -Detailed -ErrorAction Stop 2>$null)
        }
        catch {
            $cfg = @()
        }
        $netAdapters = @{}
        if (Test-NmCommandAvailable -Name Get-NetAdapter) {
            try {
                foreach ($a in @(Get-NetAdapter -ErrorAction Stop 2>$null)) {
                    $netAdapters[$a.InterfaceIndex] = $a
                }
            }
            catch {
                $netAdapters = @{}
            }
        }

        foreach ($c in $cfg) {
            $ip4 = $c.IPv4Address | Select-Object -First 1
            if (-not $ip4) { continue }

            $adapter = if ($netAdapters.ContainsKey($c.InterfaceIndex)) { $netAdapters[$c.InterfaceIndex] } else { $null }
            $desc = [string]$(if ($adapter) { $adapter.InterfaceDescription } else { $c.InterfaceDescription })
            $score = 0
            $flags = New-Object System.Collections.Generic.List[string]

            if ($adapter -and $adapter.Status -eq 'Up') { $score += 40; $flags.Add('Up') | Out-Null }
            if ($c.IPv4DefaultGateway) { $score += 15; $flags.Add('HasGateway') | Out-Null }
            if ($adapter -and $adapter.HardwareInterface) { $score += 20; $flags.Add('Physical') | Out-Null }
            if ($desc -match 'Wireless|Wi-?Fi') { $score -= 10; $flags.Add('Wireless') | Out-Null }
            if ($desc -match 'Virtual|Hyper-V|VMware|Loopback|Bluetooth|TAP|TUN|VPN|WireGuard|ZeroTier|Tailscale|AnyConnect') { $score -= 60; $flags.Add('VirtualOrVPN') | Out-Null }

            $items.Add([pscustomobject]@{
                InterfaceAlias      = $c.InterfaceAlias
                InterfaceIndex      = $c.InterfaceIndex
                InterfaceDescription= $desc
                IPv4Address         = $ip4.IPAddress
                PrefixLength        = [int]$ip4.PrefixLength
                DefaultGateway      = [string]($c.IPv4DefaultGateway.NextHop | Select-Object -First 1)
                DnsServers          = @($c.DNSServer.ServerAddresses)
                MacAddress          = if ($adapter) { $adapter.MacAddress } else { '' }
                LinkSpeed           = if ($adapter) { [string]$adapter.LinkSpeed } else { '' }
                Status              = if ($adapter) { [string]$adapter.Status } else { '' }
                Score               = $score
                Flags               = ($flags.ToArray() -join ', ')
            }) | Out-Null
        }
    }

    if ($items.Count -eq 0) {
        $legacy = @()
        try {
            $legacy = @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop)
        }
        catch {
            $legacy = @()
        }
        foreach ($c in $legacy) {
            $ip = @($c.IPAddress | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
            if (-not $ip) { continue }
            $mask = @($c.IPSubnet | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
            $prefix = if ($mask) { ([System.Net.IPAddress]::Parse($mask).GetAddressBytes() | ForEach-Object { [Convert]::ToString($_,2).PadLeft(8,'0') } | Out-String).Replace(' ','').Replace("`r",'').Replace("`n",'').Trim('0').Length } else { 24 }
            $items.Add([pscustomobject]@{
                InterfaceAlias       = $c.Description
                InterfaceIndex       = $c.InterfaceIndex
                InterfaceDescription = $c.Description
                IPv4Address          = $ip[0]
                PrefixLength         = $prefix
                DefaultGateway       = [string]($c.DefaultIPGateway | Select-Object -First 1)
                DnsServers           = @($c.DNSServerSearchOrder)
                MacAddress           = $c.MACAddress
                LinkSpeed            = ''
                Status               = 'Unknown'
                Score                = 30
                Flags                = 'Legacy'
            }) | Out-Null
        }
    }

    return @($items | Sort-Object -Property @(
        @{ Expression = 'Score'; Descending = $true },
        @{ Expression = 'InterfaceAlias'; Descending = $false }
    ))
}

function Select-NmAdapter {
    $candidates = @(Get-NmAdapterCandidates)
    if ($candidates.Count -lt 1) {
        Write-NmStatus -Level Error -Message 'No IPv4 adapters were detected.'
        return
    }

    Show-NmHeader
    Write-NmSection 'Adapter Selection'
    $i = 0
    foreach ($item in $candidates) {
        $i++
        Write-Host ('[{0}] {1,-20} {2,-15} GW={3,-15} Score={4} {5}' -f $i, $item.InterfaceAlias, $item.IPv4Address, $(if ($item.DefaultGateway) { $item.DefaultGateway } else { 'none' }), $item.Score, $item.Flags)
    }

    $pick = Read-Host ('Select adapter [1-{0}] (blank = 1)' -f $candidates.Count)
    if ([string]::IsNullOrWhiteSpace($pick)) { $pick = '1' }
    if ($pick -notmatch '^[0-9]+$') {
        Write-NmStatus -Level Warn -Message 'Invalid selection.'
        return
    }

    $index = [int]$pick - 1
    if ($index -lt 0 -or $index -ge $candidates.Count) {
        Write-NmStatus -Level Warn -Message 'Selection out of range.'
        return
    }

    $script:NmState.SelectedAdapter = $candidates[$index]
    $script:NmState.SubnetInfo = Get-NmSubnetInfo -IPAddress $script:NmState.SelectedAdapter.IPv4Address -PrefixLength $script:NmState.SelectedAdapter.PrefixLength
    Write-NmStatus -Level Success -Message ("Selected adapter: {0} ({1}/{2})" -f $script:NmState.SelectedAdapter.InterfaceAlias, $script:NmState.SelectedAdapter.IPv4Address, $script:NmState.SelectedAdapter.PrefixLength)
}

function Get-NmLocalSnapshot {
    if (-not $script:NmState.SelectedAdapter) { return $null }

    $adapter = $script:NmState.SelectedAdapter
    $wmi = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter ("InterfaceIndex={0}" -f $adapter.InterfaceIndex) -ErrorAction SilentlyContinue
    if (-not $wmi) {
        $wmi = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction SilentlyContinue | Where-Object { $_.Description -eq $adapter.InterfaceDescription } | Select-Object -First 1
    }

    return [pscustomobject]@{
        InterfaceAlias   = $adapter.InterfaceAlias
        InterfaceIndex   = $adapter.InterfaceIndex
        IPv4Address      = $adapter.IPv4Address
        PrefixLength     = $adapter.PrefixLength
        SubnetMask       = $script:NmState.SubnetInfo.SubnetMask
        DefaultGateway   = $adapter.DefaultGateway
        DnsServers       = @($adapter.DnsServers)
        MacAddress       = $adapter.MacAddress
        LinkSpeed        = $adapter.LinkSpeed
        DhcpEnabled      = if ($wmi) { [bool]$wmi.DHCPEnabled } else { $null }
        DhcpServer       = if ($wmi) { [string]$wmi.DHCPServer } else { '' }
        DnsDomain        = if ($wmi) { [string]$wmi.DNSDomain } else { '' }
    }
}

function New-NmHostRecord {
    param(
        [Parameter(Mandatory)][string]$IPAddress,
        [string]$HostName = '',
        [string]$MacAddress = '',
        [string]$Source = '',
        [bool]$RespondsToPing = $false,
        [string[]]$OpenPorts = @(),
        [string]$VendorGuess = '',
        [string]$RoleGuess = '',
        [string]$Notes = ''
    )

    [pscustomobject]@{
        IPAddress      = $IPAddress
        HostName       = $HostName
        MacAddress     = $MacAddress
        Source         = $Source
        RespondsToPing = $RespondsToPing
        OpenPorts      = @($OpenPorts)
        VendorGuess    = $VendorGuess
        RoleGuess      = $RoleGuess
        Notes          = $Notes
    }
}

function Merge-NmHostInventory {
    param([Parameter(Mandatory)][object[]]$Incoming)

    $byIp = @{}
    foreach ($item in @($script:NmState.Inventory)) {
        $byIp[$item.IPAddress] = $item
    }

    foreach ($item in @($Incoming)) {
        if ($byIp.ContainsKey($item.IPAddress)) {
            $old = $byIp[$item.IPAddress]
            $old.HostName = if ($item.HostName) { $item.HostName } else { $old.HostName }
            $old.MacAddress = if ($item.MacAddress) { $item.MacAddress } else { $old.MacAddress }
            $old.Source = (@($old.Source, $item.Source) | Where-Object { $_ } | Select-Object -Unique) -join ', '
            $old.RespondsToPing = ($old.RespondsToPing -or $item.RespondsToPing)
            $old.OpenPorts = @($old.OpenPorts + $item.OpenPorts | Where-Object { $_ } | Select-Object -Unique)
            $old.VendorGuess = if ($item.VendorGuess) { $item.VendorGuess } else { $old.VendorGuess }
            $old.RoleGuess = if ($item.RoleGuess) { $item.RoleGuess } else { $old.RoleGuess }
            $old.Notes = (@($old.Notes, $item.Notes) | Where-Object { $_ } | Select-Object -Unique) -join '; '
        }
        else {
            $byIp[$item.IPAddress] = $item
        }
    }

    $script:NmState.Inventory = @($byIp.Values | Sort-Object { ConvertTo-NmUInt32IP -IPAddress $_.IPAddress })
}

function Resolve-NmHostName {
    param([Parameter(Mandatory)][string]$IPAddress)
    try {
        return ([System.Net.Dns]::GetHostEntry($IPAddress)).HostName
    }
    catch {
        return ''
    }
}

function Get-NmNeighborInventory {
    if (-not $script:NmState.SelectedAdapter) { return @() }

    $hosts = New-Object System.Collections.Generic.List[object]
    if (Test-NmCommandAvailable -Name Get-NetNeighbor) {
        $neighbors = @(Get-NetNeighbor -InterfaceIndex $script:NmState.SelectedAdapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -match '^\d+\.' -and $_.State -notin @('Unreachable','Invalid') })
        foreach ($n in $neighbors) {
            $vendor = Resolve-NmMacVendor -MacAddress $n.LinkLayerAddress
            $hosts.Add((New-NmHostRecord -IPAddress $n.IPAddress -HostName (Resolve-NmHostName -IPAddress $n.IPAddress) -MacAddress $n.LinkLayerAddress -Source 'Neighbor' -RespondsToPing:$true -VendorGuess $vendor)) | Out-Null
        }
    }

    return $hosts.ToArray()
}

function Get-NmSweepTargets {
    param(
        [Parameter(Mandatory)]$SubnetInfo,
        [int]$MaxHosts = 510
    )

    if ($SubnetInfo.HostCount -le 0) { return @() }
    if ($SubnetInfo.HostCount -gt $MaxHosts) {
        Write-NmStatus -Level Warn -Message ("Subnet has {0} usable IPs. Sweep will be capped to {1} addresses nearest the local host." -f $SubnetInfo.HostCount, $MaxHosts)
    }

    $targets = New-Object System.Collections.Generic.List[string]
    $start = $SubnetInfo.NetworkInt + 1
    $end = $SubnetInfo.BroadcastInt - 1
    for ($value = $start; $value -le $end; $value++) {
        $targets.Add((ConvertFrom-NmUInt32IP -Value $value)) | Out-Null
        if ($targets.Count -ge $MaxHosts) { break }
    }
    return $targets.ToArray()
}

function Invoke-NmRapidSurvey {
    if (-not $script:NmState.SelectedAdapter) {
        Write-NmStatus -Level Warn -Message 'Select an adapter first.'
        return
    }

    $script:NmState.LocalSnapshot = Get-NmLocalSnapshot
    $items = New-Object System.Collections.Generic.List[object]
    $items.Add((New-NmHostRecord -IPAddress $script:NmState.LocalSnapshot.IPv4Address -HostName $env:COMPUTERNAME -MacAddress $script:NmState.LocalSnapshot.MacAddress -Source 'Local' -RespondsToPing:$true -RoleGuess 'Technician Workstation' -Notes 'Local survey host')) | Out-Null

    if ($script:NmState.LocalSnapshot.DefaultGateway) {
        $gwPing = Test-Connection -ComputerName $script:NmState.LocalSnapshot.DefaultGateway -Count 1 -Quiet -ErrorAction SilentlyContinue
        $items.Add((New-NmHostRecord -IPAddress $script:NmState.LocalSnapshot.DefaultGateway -HostName (Resolve-NmHostName -IPAddress $script:NmState.LocalSnapshot.DefaultGateway) -Source 'Gateway' -RespondsToPing:$gwPing -RoleGuess 'Gateway / Router' -Notes 'Default gateway for selected adapter')) | Out-Null
    }

    foreach ($host in (Get-NmNeighborInventory)) {
        $items.Add($host) | Out-Null
    }

    Merge-NmHostInventory -Incoming $items.ToArray()
    $script:NmState.LastSurveyAt = Get-Date
    Write-NmStatus -Level Success -Message ("Rapid survey complete. Inventory now contains {0} hosts." -f @($script:NmState.Inventory).Count)
}

function Invoke-NmSubnetSweep {
    if (-not $script:NmState.SubnetInfo) {
        Write-NmStatus -Level Warn -Message 'Select an adapter first.'
        return
    }

    $targets = @(Get-NmSweepTargets -SubnetInfo $script:NmState.SubnetInfo)
    if ($targets.Count -lt 1) {
        Write-NmStatus -Level Warn -Message 'No sweep targets available.'
        return
    }

    Write-NmSection 'Subnet Sweep'
    Write-NmStatus -Message ("Sweeping {0} IPv4 targets. This can take a little while." -f $targets.Count)
    $found = New-Object System.Collections.Generic.List[object]
    $counter = 0

    foreach ($ip in $targets) {
        $counter++
        if (($counter % 16) -eq 0 -or $counter -eq 1 -or $counter -eq $targets.Count) {
            Write-Progress -Activity 'NetworkMatrix Sweep' -Status $ip -PercentComplete (($counter / $targets.Count) * 100)
        }

        $didItWork = $false
        try {
            $didItWork = Test-Connection -ComputerName $ip -Count 1 -Quiet -ErrorAction SilentlyContinue
        }
        catch {
            $didItWork = $false
        }

        if ($didItWork) {
            $found.Add((New-NmHostRecord -IPAddress $ip -HostName (Resolve-NmHostName -IPAddress $ip) -Source 'Ping Sweep' -RespondsToPing:$true)) | Out-Null
        }
    }

    Write-Progress -Activity 'NetworkMatrix Sweep' -Completed
    Merge-NmHostInventory -Incoming $found.ToArray()
    $script:NmState.LastSurveyAt = Get-Date
    Write-NmStatus -Level Success -Message ("Sweep complete. Reachable hosts found in this pass: {0}" -f $found.Count)
}

function Test-NmTcpPort {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][int]$Port,
        [int]$TimeoutMs = 250
    )

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            return $false
        }
        $client.EndConnect($iar)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Close()
    }
}

function Resolve-NmRoleGuess {
    param([Parameter(Mandatory)]$Host)

    $ports = @($Host.OpenPorts)
    $hostName = [string]$Host.HostName
    $vendor = [string]$Host.VendorGuess

    if ($script:NmState.LocalSnapshot -and $Host.IPAddress -eq $script:NmState.LocalSnapshot.DefaultGateway) { return 'Gateway / Router' }
    if ($Host.RoleGuess -eq 'Technician Workstation') { return $Host.RoleGuess }
    if ($hostName -match 'switch|core|edge|router|firewall') { return 'Switch / Network Gear' }
    if ($vendor -match 'Cisco|Aruba|Ubiquiti|TP-Link') { return 'Switch / Network Gear' }
    if ($ports -contains 554 -and $ports -contains 80 -and $ports -contains 8000) { return 'Possible NVR / DVR' }
    if ($ports -contains 37777 -or $ports -contains 8000) { return 'Possible NVR / DVR' }
    if ($ports -contains 554) { return 'Possible Camera' }
    if ($vendor -match 'Hikvision|Dahua|Axis|Hanwha|Reolink|Uniview') { return 'Possible Camera' }
    if ($ports -contains 445 -or $ports -contains 3389) { return 'Workstation / Server' }
    return 'Unknown'
}

function Invoke-NmPortFingerprint {
    if (@($script:NmState.Inventory).Count -lt 1) {
        Write-NmStatus -Level Warn -Message 'Run a rapid survey or subnet sweep first.'
        return
    }

    $targets = @($script:NmState.Inventory | Where-Object { $_.RespondsToPing -and $_.RoleGuess -ne 'Technician Workstation' })
    if ($targets.Count -gt 64) {
        $targets = @($targets | Select-Object -First 64)
        Write-NmStatus -Level Warn -Message 'Port fingerprinting is capped to the first 64 reachable hosts for speed.'
    }

    foreach ($tmpNet in $targets) {
        $open = New-Object System.Collections.Generic.List[string]
        foreach ($port in $script:NmKnownPorts) {
            if (Test-NmTcpPort -ComputerName $tmpNet.IPAddress -Port $port) {
                $open.Add([string]$port) | Out-Null
            }
        }
        $tmpNet.OpenPorts = $open.ToArray()
        if (-not $tmpNet.VendorGuess -and $tmpNet.MacAddress) {
            $tmpNet.VendorGuess = Resolve-NmMacVendor -MacAddress $tmpNet.MacAddress
        }
        $tmpNet.RoleGuess = Resolve-NmRoleGuess -Host $tmpNet
    }

    $script:NmState.LastPortScanAt = Get-Date
    Write-NmStatus -Level Success -Message ("Port fingerprinting complete for {0} hosts." -f $targets.Count)
}

function Find-NmFreeBlock {
    param(
        [Parameter(Mandatory)]$SubnetInfo,
        [Parameter(Mandatory)][string[]]$UsedIPs,
        [Parameter(Mandatory)][int]$Size,
        [int]$StartOffset = 150
    )

    $used = @{}
    foreach ($ip in $UsedIPs) { if ($ip) { $used[$ip] = $true } }

    $start = [math]::Max(($SubnetInfo.NetworkInt + 1 + $StartOffset), ($SubnetInfo.NetworkInt + 1))
    $end = $SubnetInfo.BroadcastInt - 1
    for ($base = $start; $base -le $end; $base++) {
        $ok = $true
        for ($i = 0; $i -lt $Size; $i++) {
            $current = $base + $i
            if ($current -gt $end) { $ok = $false; break }
            $tmp = ConvertFrom-NmUInt32IP -Value $current
            if ($used.ContainsKey($tmp)) { $ok = $false; break }
        }
        if ($ok) {
            return [pscustomobject]@{
                Start = ConvertFrom-NmUInt32IP -Value $base
                End   = ConvertFrom-NmUInt32IP -Value ($base + $Size - 1)
                Size  = $Size
            }
        }
    }
    return $null
}

function Get-NmDeploymentBriefText {
    if (-not $script:NmState.LocalSnapshot) {
        $script:NmState.LocalSnapshot = Get-NmLocalSnapshot
    }
    if (-not $script:NmState.LocalSnapshot) {
        return 'No adapter survey is loaded yet.'
    }

    $inventory = @($script:NmState.Inventory)
    $used = @($inventory.IPAddress + $script:NmState.LocalSnapshot.IPv4Address + $script:NmState.LocalSnapshot.DefaultGateway | Where-Object { $_ } | Select-Object -Unique)
    $cameraBlock = Find-NmFreeBlock -SubnetInfo $script:NmState.SubnetInfo -UsedIPs $used -Size 16 -StartOffset 150
    $recorderBlock = Find-NmFreeBlock -SubnetInfo $script:NmState.SubnetInfo -UsedIPs $used -Size 8 -StartOffset 220
    $likelyCameras = @($inventory | Where-Object { $_.RoleGuess -eq 'Possible Camera' })
    $likelyRecorders = @($inventory | Where-Object { $_.RoleGuess -eq 'Possible NVR / DVR' })
    $networkGear = @($inventory | Where-Object { $_.RoleGuess -eq 'Switch / Network Gear' -or $_.RoleGuess -eq 'Gateway / Router' })

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('NetworkMatrix Deployment Brief') | Out-Null
    $lines.Add(('Generated: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))) | Out-Null
    $lines.Add(('Host: {0}' -f $env:COMPUTERNAME)) | Out-Null
    $lines.Add(('Admin: {0}' -f $(if (Test-NmIsAdmin) { 'Yes' } else { 'No' }))) | Out-Null
    $lines.Add(('=' * 78)) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Surveyed Network') | Out-Null
    $lines.Add((' Adapter:      {0}' -f $script:NmState.LocalSnapshot.InterfaceAlias)) | Out-Null
    $lines.Add((' IP Address:   {0}/{1}' -f $script:NmState.LocalSnapshot.IPv4Address, $script:NmState.LocalSnapshot.PrefixLength)) | Out-Null
    $lines.Add((' Subnet Mask:  {0}' -f $script:NmState.LocalSnapshot.SubnetMask)) | Out-Null
    $lines.Add((' Gateway:      {0}' -f $(if ($script:NmState.LocalSnapshot.DefaultGateway) { $script:NmState.LocalSnapshot.DefaultGateway } else { 'none detected' }))) | Out-Null
    $lines.Add((' DNS:          {0}' -f $(if (@($script:NmState.LocalSnapshot.DnsServers).Count -gt 0) { @($script:NmState.LocalSnapshot.DnsServers) -join ', ' } else { 'none detected' }))) | Out-Null
    $lines.Add((' DHCP:         {0}' -f $(if ($script:NmState.LocalSnapshot.DhcpEnabled -eq $true) { 'Enabled' } elseif ($script:NmState.LocalSnapshot.DhcpEnabled -eq $false) { 'Disabled' } else { 'Unknown' }))) | Out-Null
    $lines.Add((' DHCP Server:  {0}' -f $(if ($script:NmState.LocalSnapshot.DhcpServer) { $script:NmState.LocalSnapshot.DhcpServer } else { 'unknown' }))) | Out-Null
    $lines.Add((' Network:      {0}' -f $script:NmState.SubnetInfo.NetworkAddress)) | Out-Null
    $lines.Add((' Broadcast:    {0}' -f $script:NmState.SubnetInfo.BroadcastAddress)) | Out-Null
    $lines.Add((' Usable Hosts: {0}' -f $script:NmState.SubnetInfo.HostCount)) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add(('Discovered Hosts: {0}' -f $inventory.Count)) | Out-Null
    $lines.Add((' Possible Cameras: {0}' -f $likelyCameras.Count)) | Out-Null
    $lines.Add((' Possible NVR/DVR: {0}' -f $likelyRecorders.Count)) | Out-Null
    $lines.Add((' Network Gear: {0}' -f $networkGear.Count)) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Suggested Static IP Planning') | Out-Null
    $lines.Add((' Camera Block:   {0}' -f $(if ($cameraBlock) { '{0} -> {1} ({2} IPs)' -f $cameraBlock.Start, $cameraBlock.End, $cameraBlock.Size } else { 'No clean 16-IP block found in scanned range' }))) | Out-Null
    $lines.Add((' Recorder Block: {0}' -f $(if ($recorderBlock) { '{0} -> {1} ({2} IPs)' -f $recorderBlock.Start, $recorderBlock.End, $recorderBlock.Size } else { 'No clean 8-IP block found in scanned range' }))) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Implementation Notes') | Out-Null
    $lines.Add(' 1. Confirm the gateway and DHCP server before assigning static camera addresses.') | Out-Null
    $lines.Add(' 2. Keep cameras and recorder in the same subnet unless routing/VLAN policy requires otherwise.') | Out-Null
    $lines.Add(' 3. Validate recorder-relevant ports: 80/443, 554, 8000, 37777, 8080, 8899.') | Out-Null
    $lines.Add(' 4. Confirm PoE switch uplinks and available ports before mounting cameras.') | Out-Null
    $lines.Add(' 5. Reserve or document every static address used so the site can be serviced later.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Known Hosts') | Out-Null
    foreach ($item in $inventory) {
        $lines.Add((' - {0,-15} {1,-20} {2,-18} Ports={3}' -f $item.IPAddress, $(if ($item.RoleGuess) { $item.RoleGuess } else { 'Unknown' }), $(if ($item.VendorGuess) { $item.VendorGuess } else { '-' }), $(if (@($item.OpenPorts).Count -gt 0) { @($item.OpenPorts) -join ',' } else { '-' }))) | Out-Null
    }

    return ($lines.ToArray() -join [Environment]::NewLine)
}

function Get-NmMermaidMap {
    $inventory = @($script:NmState.Inventory)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('```mermaid') | Out-Null
    $lines.Add('graph LR') | Out-Null
    $lines.Add(('    local["Technician`n{0}"]' -f $script:NmState.LocalSnapshot.IPv4Address)) | Out-Null
    if ($script:NmState.LocalSnapshot.DefaultGateway) {
        $lines.Add(('    gateway["Gateway`n{0}"]' -f $script:NmState.LocalSnapshot.DefaultGateway)) | Out-Null
        $lines.Add('    local --> gateway') | Out-Null
    }

    $i = 0
    foreach ($host in $inventory) {
        if ($host.RoleGuess -eq 'Technician Workstation') { continue }
        $i++
        $node = 'node{0}' -f $i
        $label = '{0}\n{1}\n{2}' -f $(if ($host.RoleGuess) { $host.RoleGuess } else { 'Unknown' }), $host.IPAddress, $(if ($host.VendorGuess) { $host.VendorGuess } else { '-' })
        $lines.Add(('    {0}["{1}"]' -f $node, $label.Replace('"',"'"))) | Out-Null
        if ($script:NmState.LocalSnapshot.DefaultGateway -and $host.IPAddress -ne $script:NmState.LocalSnapshot.DefaultGateway) {
            $lines.Add(('    gateway --> {0}' -f $node)) | Out-Null
        }
        else {
            $lines.Add(('    local --> {0}' -f $node)) | Out-Null
        }
    }

    $lines.Add('```') | Out-Null
    return ($lines.ToArray() -join [Environment]::NewLine)
}

function Show-NmCurrentMap {
    Show-NmHeader
    Write-NmSection 'Current Network Map'
    if (-not $script:NmState.LocalSnapshot) {
        Write-NmStatus -Level Warn -Message 'No survey data loaded yet.'
        return
    }

    Write-Host ('[Technician] {0} on {1}' -f $script:NmState.LocalSnapshot.IPv4Address, $script:NmState.LocalSnapshot.InterfaceAlias) -ForegroundColor White
    if ($script:NmState.LocalSnapshot.DefaultGateway) {
        Write-Host ('  +-- [Gateway] {0}' -f $script:NmState.LocalSnapshot.DefaultGateway) -ForegroundColor Yellow
    }

    foreach ($item in @($script:NmState.Inventory | Where-Object { $_.RoleGuess -ne 'Technician Workstation' })) {
        $ports = if (@($item.OpenPorts).Count -gt 0) { @($item.OpenPorts) -join ',' } else { '-' }
        Write-Host ('      +-- [{0}] {1,-15} {2,-15} Ports={3}' -f $(if ($item.RoleGuess) { $item.RoleGuess } else { 'Unknown' }), $item.IPAddress, $(if ($item.VendorGuess) { $item.VendorGuess } else { '-' }), $ports) -ForegroundColor Gray
    }
}

function Export-NmSurvey {
    if (-not $script:NmState.LocalSnapshot) {
        Write-NmStatus -Level Warn -Message 'No survey data loaded yet.'
        return
    }

    $brief = Get-NmDeploymentBriefText
    $mermaid = Get-NmMermaidMap
    $inventory = @($script:NmState.Inventory | ForEach-Object {
        [pscustomobject]@{
            IPAddress      = $_.IPAddress
            HostName       = $_.HostName
            MacAddress     = $_.MacAddress
            Source         = $_.Source
            RespondsToPing = $_.RespondsToPing
            OpenPorts      = @($_.OpenPorts) -join ','
            VendorGuess    = $_.VendorGuess
            RoleGuess      = $_.RoleGuess
            Notes          = $_.Notes
        }
    })

    $textPath = Save-NmTextReport -Prefix 'NetworkMatrix_Brief' -Text ($brief + [Environment]::NewLine + [Environment]::NewLine + $mermaid)
    $jsonPath = Save-NmJsonReport -Prefix 'NetworkMatrix_Data' -Object ([pscustomobject]@{
        Session    = $script:NmState.SessionName
        Adapter    = $script:NmState.SelectedAdapter
        Local      = $script:NmState.LocalSnapshot
        Subnet     = $script:NmState.SubnetInfo
        Inventory  = $inventory
        Generated  = Get-Date
    })
    $csvPath = Save-NmCsvReport -Prefix 'NetworkMatrix_Inventory' -Object $inventory
    $mdPath = Save-NmTextReport -Prefix 'NetworkMatrix_Map' -Text ('# NetworkMatrix Map' + [Environment]::NewLine + [Environment]::NewLine + $mermaid) -Extension 'md'

    $script:NmState.LastBrief = $brief
    Write-NmStatus -Level Success -Message ('Exported survey package: {0}, {1}, {2}, {3}' -f $textPath, $jsonPath, $csvPath, $mdPath)
}

function Show-NmMenu {
    while ($true) {
        Show-NmHeader
        Write-Host ''
        Write-Host '[1] Select target adapter' -ForegroundColor Cyan
        Write-Host '[2] Run rapid site survey' -ForegroundColor Cyan
        Write-Host '[3] Run subnet sweep' -ForegroundColor Cyan
        Write-Host '[4] Fingerprint likely camera/NVR ports' -ForegroundColor Cyan
        Write-Host '[5] Show current network map' -ForegroundColor Cyan
        Write-Host '[6] Build deployment brief' -ForegroundColor Cyan
        Write-Host '[7] Export survey package' -ForegroundColor Cyan
        Write-Host '[8] Exit' -ForegroundColor Cyan
        Write-Host ''

        $pick = Read-Host 'Choose an action'
        switch ($pick) {
            '1' { Select-NmAdapter; Pause-Nm }
            '2' { Invoke-NmRapidSurvey; Pause-Nm }
            '3' { Invoke-NmSubnetSweep; Pause-Nm }
            '4' { Invoke-NmPortFingerprint; Pause-Nm }
            '5' { Show-NmCurrentMap; Pause-Nm }
            '6' {
                Show-NmHeader
                Write-NmSection 'Deployment Brief'
                $brief = Get-NmDeploymentBriefText
                $script:NmState.LastBrief = $brief
                $brief
                Pause-Nm
            }
            '7' { Export-NmSurvey; Pause-Nm }
            '8' { return }
            default {
                Write-NmStatus -Level Warn -Message 'Unknown selection.'
                Pause-Nm
            }
        }
    }
}

Initialize-NmReportSession -Name $ReportName -NonInteractive:$NonInteractive

if ($NonInteractive) {
    if (-not $script:NmState.SelectedAdapter) {
        $picked = @(Get-NmAdapterCandidates | Select-Object -First 1)
        if ($picked.Count -gt 0) {
            $script:NmState.SelectedAdapter = $picked[0]
            $script:NmState.SubnetInfo = Get-NmSubnetInfo -IPAddress $script:NmState.SelectedAdapter.IPv4Address -PrefixLength $script:NmState.SelectedAdapter.PrefixLength
        }
    }
    if (-not $script:NmState.SelectedAdapter) {
        Write-NmStatus -Level Warn -Message 'No usable IPv4 adapter could be selected in non-interactive mode.'
        return
    }
    Invoke-NmRapidSurvey
    Invoke-NmPortFingerprint
    Export-NmSurvey
    return
}

Show-NmMenu
