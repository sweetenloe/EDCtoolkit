#Requires -Version 5.1

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Continue'

$ToolkitBannerRaw = @"
        '+--------------------------------------------------+',
        '|                                                  |',
        '|   ********************************************   |',
        '|   *                                          *   |',
        '|   *    _       _   _ _|_  _  ._  |  _   _    *   |',
        '|   *   _> \/\/ (/_ (/_ |_ (/_ | | | (_) (/_   *   |',
        '|   *                                          *   |',
        '|   *                                          *   |',
        '|   ********************************************   |',
        '|                                                  |',
        '|           0 calories | Sugar Substitute          |',
        '|                                                  |',
        '+--------------------------------------------------+'
                  Portable Field Technician Toolkit
"@

$ToolkitBanner = @($ToolkitBannerRaw -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

$Script:ToolkitName = 'EDCtoolkit'
$Script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$Script:ReportBaseRoot = Join-Path -Path $Script:ScriptRoot -ChildPath 'EDC_Reports'
$Script:ReportSessionName = $null
$Script:ReportRoot = $Script:ReportBaseRoot
$Script:LastReports = New-Object System.Collections.Generic.List[string]

if (-not (Test-Path -Path $Script:ReportBaseRoot)) {
    New-Item -Path $Script:ReportBaseRoot -ItemType Directory -Force | Out-Null
}

function Get-Timestamp {
    Get-Date -Format 'yyyyMMdd_HHmmss'
}

function Test-IsAdmin {
    try {
        $current = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($current)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Test-CommandAvailable {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Convert-ToSafeFolderName {
    param([Parameter(Mandatory)][string]$Name)
    $safe = $Name.Trim()
    $safe = $safe -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace '\s+', '_'
    $safe = $safe.Trim('.')
    return $safe
}

function Convert-FromDmtfDateTime {
    param(
        [AllowNull()]$Value,
        [switch]$Quiet
    )

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return [datetime]$Value }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($text)
    }
    catch {
        if (-not $Quiet) {
            Write-Status -Level Warn -Message "Unable to parse WMI datetime value: '$text'"
        }
        return $null
    }
}

function Show-Header {
    Clear-Host
    foreach ($line in $ToolkitBanner) {
        Write-Host $line -ForegroundColor Magenta
    }
    Write-Host ('=' * 74) -ForegroundColor DarkGray
    Write-Host (' Toolkit: {0}  |  Reports: {1}' -f $Script:ToolkitName, $Script:ReportRoot) -ForegroundColor Gray
    Write-Host (' Session: {0}  |  Admin: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if (Test-IsAdmin) { 'Yes' } else { 'No' })) -ForegroundColor Gray
    Write-Host ('=' * 74) -ForegroundColor DarkGray
}

function Initialize-ReportSession {
    while ($true) {
        Show-Header
        Write-Section 'Report Session Setup'
        $rawName = Read-Host 'Name The Report'
        if ([string]::IsNullOrWhiteSpace($rawName)) {
            $rawName = 'Report_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
        }
        $safeName = Convert-ToSafeFolderName -Name $rawName
        if ([string]::IsNullOrWhiteSpace($safeName)) {
            Write-Status -Level Warn -Message 'Report name is invalid after sanitization. Please enter another name.'
            Pause-Toolkit
            continue
        }
        $target = Join-Path -Path $Script:ReportBaseRoot -ChildPath $safeName
        try {
            if (-not (Test-Path -Path $target)) { New-Item -Path $target -ItemType Directory -Force | Out-Null }
            $Script:ReportSessionName = $safeName
            $Script:ReportRoot = $target
            Write-Status -Level Success -Message "Reports will save to: $Script:ReportRoot"
            Pause-Toolkit
            return
        }
        catch {
            Write-Status -Level Error -Message ("Unable to create report folder '{0}': {1}" -f $safeName, $_.Exception.Message)
            Pause-Toolkit
        }
    }
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)
    Write-Host ''
    Write-Host ('--- {0} ---' -f $Title) -ForegroundColor Yellow
}

function Write-Status {
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

function Pause-Toolkit {
    Read-Host 'Press Enter to continue'
}

function Confirm-Action {
    param([Parameter(Mandatory)][string]$Prompt)
    $choice = Read-Host "$Prompt (Y/N)"
    return $choice -match '^(y|yes)$'
}

function Show-AdminHint {
    param([Parameter(Mandatory)][string]$Action)
    if (-not (Test-IsAdmin)) {
        Write-Status -Level Warn -Message "Action '$Action' may require Run as Administrator for full results."
    }
}

function New-ReportPath {
    param([Parameter(Mandatory)][string]$Prefix)
    $safePrefix = ($Prefix -replace '[^a-zA-Z0-9_-]', '_')
    $name = '{0}_{1}.txt' -f $safePrefix, (Get-Timestamp)
    return Join-Path -Path $Script:ReportRoot -ChildPath $name
}

function Save-TextReport {
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Text
    )
    $path = New-ReportPath -Prefix $Prefix
    $header = @(
        ('Toolkit: {0}' -f $Script:ToolkitName),
        ('Generated: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
        ('Host: {0}' -f $env:COMPUTERNAME),
        ('User: {0}' -f $env:USERNAME),
        ('=' * 70),
        ''
    ) -join [Environment]::NewLine

    ($header + $Text) | Set-Content -Path $path -Encoding UTF8
    $Script:LastReports.Add($path)
    Write-Status -Level Success -Message "Report saved: $path"
    return $path
}

function Save-ObjectReport {
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)]$Object,
        [string]$Format = 'Table'
    )
    $body = if ($null -eq $Object) {
        'No data returned.'
    }
    elseif ($Format -eq 'List') {
        $Object | Format-List * | Out-String -Width 4096
    }
    else {
        $Object | Format-Table -AutoSize | Out-String -Width 4096
    }
    Save-TextReport -Prefix $Prefix -Text $body | Out-Null
}

function Show-AndSaveObject {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)]$Object,
        [string]$Format = 'Table'
    )
    Write-Section $Title
    if ($null -eq $Object -or (($Object -is [System.Collections.IEnumerable]) -and -not ($Object | Select-Object -First 1))) {
        Write-Status -Level Warn -Message 'No data returned.'
        Save-TextReport -Prefix $Prefix -Text 'No data returned.' | Out-Null
        return
    }

    if ($Format -eq 'List') {
        $Object | Format-List *
    }
    else {
        $Object | Format-Table -AutoSize
    }
    Save-ObjectReport -Prefix $Prefix -Object $Object -Format $Format
}

function Show-AndSaveText {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Text
    )
    Write-Section $Title
    $Text
    Save-TextReport -Prefix $Prefix -Text $Text | Out-Null
}

function Invoke-Safe {
    param(
        [Parameter(Mandatory)][string]$ActionName,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )
    try {
        & $ScriptBlock
    }
    catch {
        Write-Status -Level Error -Message ("{0} failed: {1}" -f $ActionName, $_.Exception.Message)
    }
}

function Get-UninstallRegistryEntries {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $items = foreach ($path in $paths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
    }

    $items | Sort-Object DisplayName -Unique
}

function Invoke-FullSystemAudit {
    Write-Section 'Full System Audit'
    Write-Status -Message 'Collecting core system data. This can take a minute.'

    $computer = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue
    $ram = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
    $disk = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
    $lastBoot = Convert-FromDmtfDateTime -Value $os.LastBootUpTime -Quiet
    $uptime = if ($lastBoot) { (Get-Date) - $lastBoot } else { $null }

    $report = @()
    $report += "Host: $($env:COMPUTERNAME)"
    $report += "User: $($env:USERNAME)"
    $report += "Collected: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += ('=' * 70)
    $report += ''
    $report += '[System]'
    $report += "Manufacturer: $($computer.Manufacturer)"
    $report += "Model: $($computer.Model)"
    $report += "Serial: $($bios.SerialNumber)"
    $report += "Domain: $($computer.Domain)"
    $report += ''
    $report += '[OS]'
    $report += "Caption: $($os.Caption)"
    $report += "Version: $($os.Version)"
    $report += "Build: $($os.BuildNumber)"
    $report += "Install Date: $($os.InstallDate)"
    if ($uptime) { $report += "Uptime: $([math]::Round($uptime.TotalHours,2)) hours" }
    $report += ''
    $report += '[CPU]'
    foreach ($c in $cpu) {
        $report += "Name: $($c.Name)"
        $report += "Cores/Logical: $($c.NumberOfCores)/$($c.NumberOfLogicalProcessors)"
        $report += "MaxClockMHz: $($c.MaxClockSpeed)"
    }
    $report += ''
    $report += '[RAM]'
    $totalRam = ($ram | Measure-Object -Property Capacity -Sum).Sum
    $report += "Total RAM: $([math]::Round($totalRam / 1GB,2)) GB"
    foreach ($m in $ram) {
        $report += "Slot: $($m.DeviceLocator) | SizeGB: $([math]::Round($m.Capacity/1GB,2)) | SpeedMHz: $($m.Speed)"
    }
    $report += ''
    $report += '[Disks]'
    foreach ($d in $disk) {
        $report += "Drive $($d.DeviceID) | SizeGB: $([math]::Round($d.Size/1GB,2)) | FreeGB: $([math]::Round($d.FreeSpace/1GB,2))"
    }

    $text = $report -join [Environment]::NewLine
    Show-AndSaveText -Title 'Full System Audit (Summary)' -Prefix 'System_FullAudit' -Text $text
}

function Get-BasicHardwareInfo {
    $data = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue |
        Select-Object Manufacturer, Model, SystemType, TotalPhysicalMemory, Domain

    if ($data) {
        $data.TotalPhysicalMemory = '{0:N2} GB' -f ($data.TotalPhysicalMemory / 1GB)
    }

    Show-AndSaveObject -Title 'Basic Hardware Info' -Prefix 'System_BasicHardware' -Object $data -Format List
}

function Get-OSInfo {
    $data = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue |
        Select-Object Caption, Version, BuildNumber, OSArchitecture, InstallDate, LastBootUpTime

    Show-AndSaveObject -Title 'OS Info' -Prefix 'System_OSInfo' -Object $data -Format List
}

function Get-CPUInfo {
    $data = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue |
        Select-Object Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed

    Show-AndSaveObject -Title 'CPU Info' -Prefix 'System_CPUInfo' -Object $data
}

function Get-RAMInfo {
    $data = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue |
        Select-Object DeviceLocator, Manufacturer, Capacity, Speed, SerialNumber

    if ($data) {
        $data = $data | ForEach-Object {
            [pscustomobject]@{
                DeviceLocator = $_.DeviceLocator
                Manufacturer  = $_.Manufacturer
                CapacityGB     = [math]::Round(($_.Capacity / 1GB),2)
                SpeedMHz       = $_.Speed
                SerialNumber   = $_.SerialNumber
            }
        }
    }

    Show-AndSaveObject -Title 'RAM Info' -Prefix 'System_RAMInfo' -Object $data
}

function Get-DiskInfo {
    $data = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue |
        Select-Object DeviceID, VolumeName,
            @{Name='SizeGB';Expression={[math]::Round($_.Size/1GB,2)}},
            @{Name='FreeGB';Expression={[math]::Round($_.FreeSpace/1GB,2)}},
            @{Name='UsedGB';Expression={[math]::Round(($_.Size-$_.FreeSpace)/1GB,2)}}

    Show-AndSaveObject -Title 'Disk Info' -Prefix 'System_DiskInfo' -Object $data
}

function Get-TopProcesses {
    $topCpu = Get-Process -ErrorAction SilentlyContinue | Sort-Object CPU -Descending | Select-Object -First 15 Name,Id,CPU,WS
    $topMem = Get-Process -ErrorAction SilentlyContinue | Sort-Object WS -Descending | Select-Object -First 15 Name,Id,CPU,
        @{Name='MemoryMB';Expression={[math]::Round($_.WS/1MB,2)}}

    Write-Section 'Top Processes by CPU'
    $topCpu | Format-Table -AutoSize
    Write-Section 'Top Processes by RAM'
    $topMem | Format-Table -AutoSize

    $text = @()
    $text += 'Top Processes by CPU'
    $text += ($topCpu | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Top Processes by RAM'
    $text += ($topMem | Format-Table -AutoSize | Out-String -Width 4096)
    Save-TextReport -Prefix 'System_TopProcesses' -Text ($text -join [Environment]::NewLine) | Out-Null
}

function Get-InstalledSoftwareInventory {
    $data = Get-UninstallRegistryEntries
    Show-AndSaveObject -Title 'Installed Software Inventory' -Prefix 'System_InstalledSoftware' -Object $data
}

function Get-HotfixInventory {
    $data = Get-HotFix -ErrorAction SilentlyContinue | Select-Object HotFixID, InstalledOn, Description
    Show-AndSaveObject -Title 'Installed Windows Updates / Hotfixes' -Prefix 'System_Hotfixes' -Object $data
}

function Get-EventLogErrors24h {
    $start = (Get-Date).AddHours(-24)
    $events = Get-WinEvent -FilterHashtable @{ LogName='System'; Level=2; StartTime=$start } -ErrorAction SilentlyContinue |
        Select-Object -First 200 TimeCreated, Id, ProviderName, Message

    Show-AndSaveObject -Title 'System Event Log Errors (Last 24 Hours)' -Prefix 'System_EventErrors24h' -Object $events
}

function Get-StartupEntriesAudit {
    $wmi = Get-CimInstance Win32_StartupCommand -ErrorAction SilentlyContinue |
        Select-Object Name, Command, Location, User

    $runKeys = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce'
    )

    $regEntries = foreach ($key in $runKeys) {
        $item = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if ($item) {
            foreach ($prop in $item.PSObject.Properties) {
                if ($prop.Name -notmatch '^PS') {
                    [pscustomobject]@{
                        Source   = $key
                        Name     = $prop.Name
                        Command  = [string]$prop.Value
                    }
                }
            }
        }
    }

    Write-Section 'Startup Entries (WMI)'
    if ($wmi) { $wmi | Format-Table -AutoSize } else { Write-Status -Level Warn -Message 'No WMI startup entries found.' }

    Write-Section 'Startup Entries (Registry Run Keys)'
    if ($regEntries) { $regEntries | Format-Table -AutoSize } else { Write-Status -Level Warn -Message 'No Run key entries found.' }

    $text = @()
    $text += 'Startup Entries (WMI)'
    $text += ($wmi | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Startup Entries (Registry Run Keys)'
    $text += ($regEntries | Format-Table -AutoSize | Out-String -Width 4096)
    Save-TextReport -Prefix 'System_StartupAudit' -Text ($text -join [Environment]::NewLine) | Out-Null
}

function Get-ScheduledTasksSummary {
    if (Test-CommandAvailable -Name Get-ScheduledTask) {
        $data = Get-ScheduledTask -ErrorAction SilentlyContinue |
            Select-Object TaskName, TaskPath, State,
                @{Name='LastRunTime';Expression={$_.LastRunTime}},
                @{Name='NextRunTime';Expression={$_.NextRunTime}}
        Show-AndSaveObject -Title 'Scheduled Tasks Summary' -Prefix 'System_ScheduledTasks' -Object $data
    }
    else {
        $text = schtasks /query /fo LIST /v 2>&1 | Out-String
        Show-AndSaveText -Title 'Scheduled Tasks Summary' -Prefix 'System_ScheduledTasks' -Text $text
    }
}

function Get-EnvironmentInfo {
    $vars = Get-ChildItem Env: | Sort-Object Name
    Show-AndSaveObject -Title 'Environment Info' -Prefix 'System_Environment' -Object $vars
}

function Get-SystemUptime {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if (-not $os -or -not $os.LastBootUpTime) {
        Show-AndSaveText -Title 'System Uptime' -Prefix 'System_Uptime' -Text 'Unable to determine system uptime.'
        return
    }

    $lastBoot = Convert-FromDmtfDateTime -Value $os.LastBootUpTime
    if (-not $lastBoot) {
        Show-AndSaveText -Title 'System Uptime' -Prefix 'System_Uptime' -Text 'Unable to determine system uptime (invalid boot timestamp format).'
        return
    }

    $uptime = (Get-Date) - $lastBoot
    $text = @(
        "Last Boot: $lastBoot",
        ('Uptime: {0} days {1} hours {2} minutes' -f $uptime.Days, $uptime.Hours, $uptime.Minutes)
    ) -join [Environment]::NewLine

    Show-AndSaveText -Title 'System Uptime' -Prefix 'System_Uptime' -Text $text
}
function Get-IPConfiguration {
    if (Test-CommandAvailable -Name Get-NetIPConfiguration) {
        $data = Get-NetIPConfiguration -ErrorAction SilentlyContinue |
            Select-Object InterfaceAlias, InterfaceDescription, IPv4Address, IPv6Address, IPv4DefaultGateway, DNSServer
        Show-AndSaveObject -Title 'IP Configuration' -Prefix 'Network_IPConfig' -Object $data
    }
    else {
        $text = ipconfig /all | Out-String
        Show-AndSaveText -Title 'IP Configuration' -Prefix 'Network_IPConfig' -Text $text
    }
}

function Get-AdapterSummary {
    if (Test-CommandAvailable -Name Get-NetAdapter) {
        $data = Get-NetAdapter -ErrorAction SilentlyContinue |
            Select-Object Name, InterfaceDescription, Status, LinkSpeed, MacAddress
        Show-AndSaveObject -Title 'Network Adapter Summary' -Prefix 'Network_AdapterSummary' -Object $data
    }
    else {
        $text = netsh interface show interface | Out-String
        Show-AndSaveText -Title 'Network Adapter Summary' -Prefix 'Network_AdapterSummary' -Text $text
    }
}

function Get-ArpTable {
    $text = arp -a | Out-String
    Show-AndSaveText -Title 'ARP Table' -Prefix 'Network_ARP' -Text $text
}

function Get-NeighborTable {
    if (Test-CommandAvailable -Name Get-NetNeighbor) {
        $data = Get-NetNeighbor -ErrorAction SilentlyContinue |
            Select-Object ifIndex, IPAddress, LinkLayerAddress, State
        Show-AndSaveObject -Title 'Neighbor Table' -Prefix 'Network_Neighbors' -Object $data
    }
    else {
        Write-Status -Level Warn -Message 'Get-NetNeighbor is unavailable on this system.'
    }
}

function Get-RouteTable {
    if (Test-CommandAvailable -Name Get-NetRoute) {
        $data = Get-NetRoute -ErrorAction SilentlyContinue |
            Select-Object DestinationPrefix, NextHop, RouteMetric, InterfaceAlias, AddressFamily
        Show-AndSaveObject -Title 'Route Table' -Prefix 'Network_Routes' -Object $data
    }
    else {
        $text = route print | Out-String
        Show-AndSaveText -Title 'Route Table' -Prefix 'Network_Routes' -Text $text
    }
}

function Invoke-DnsResolutionTest {
    $name = Read-Host 'Enter DNS name to resolve'
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Status -Level Warn -Message 'DNS name cannot be empty.'
        return
    }

    if (Test-CommandAvailable -Name Resolve-DnsName) {
        $data = Resolve-DnsName -Name $name -ErrorAction SilentlyContinue | Select-Object Name, Type, IPAddress, NameHost
        Show-AndSaveObject -Title "DNS Resolution Test: $name" -Prefix 'Network_DNSResolution' -Object $data
    }
    else {
        $text = nslookup $name | Out-String
        Show-AndSaveText -Title "DNS Resolution Test: $name" -Prefix 'Network_DNSResolution' -Text $text
    }
}

function Invoke-PingTest {
    $target = Read-Host 'Enter host/IP to ping'
    if ([string]::IsNullOrWhiteSpace($target)) {
        Write-Status -Level Warn -Message 'Ping target cannot be empty.'
        return
    }

    $text = Test-Connection -ComputerName $target -Count 4 -ErrorAction SilentlyContinue | Out-String
    if ([string]::IsNullOrWhiteSpace($text)) {
        $text = ping $target | Out-String
    }
    Show-AndSaveText -Title "Ping Test: $target" -Prefix 'Network_PingTest' -Text $text
}

function Get-ListeningPorts {
    if (Test-CommandAvailable -Name Get-NetTCPConnection) {
        $data = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue |
            Select-Object LocalAddress, LocalPort, OwningProcess
        Show-AndSaveObject -Title 'Open / Listening Ports' -Prefix 'Network_ListeningPorts' -Object $data
    }
    else {
        $text = netstat -ano | Select-String 'LISTENING' | Out-String
        Show-AndSaveText -Title 'Open / Listening Ports' -Prefix 'Network_ListeningPorts' -Text $text
    }
}

function Get-ActiveTcpConnections {
    if (Test-CommandAvailable -Name Get-NetTCPConnection) {
        $data = Get-NetTCPConnection -ErrorAction SilentlyContinue |
            Select-Object State, LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess
        Show-AndSaveObject -Title 'Active TCP Connections' -Prefix 'Network_ActiveTCP' -Object $data
    }
    else {
        $text = netstat -ano | Out-String
        Show-AndSaveText -Title 'Active TCP Connections' -Prefix 'Network_ActiveTCP' -Text $text
    }
}

function Invoke-NetworkReset {
    Show-AdminHint -Action 'Network Reset'
    if (-not (Confirm-Action -Prompt 'Reset Winsock and IP stack? This is disruptive')) {
        Write-Status -Message 'Network reset canceled.'
        return
    }

    Invoke-Safe -ActionName 'netsh winsock reset' -ScriptBlock { netsh winsock reset | Out-Host }
    Invoke-Safe -ActionName 'netsh int ip reset' -ScriptBlock { netsh int ip reset | Out-Host }
    Save-TextReport -Prefix 'Network_Reset' -Text 'Network reset commands executed (winsock reset + int ip reset).' | Out-Null
}

function Get-LocalSharesAndMappedDrives {
    $shares = if (Test-CommandAvailable -Name Get-SmbShare) {
        Get-SmbShare -ErrorAction SilentlyContinue | Select-Object Name, Path, Description
    }

    $mapped = Get-PSDrive -PSProvider FileSystem | Select-Object Name, Root,
        @{Name='FreeGB';Expression={[math]::Round($_.Free/1GB,2)}},
        @{Name='UsedGB';Expression={[math]::Round($_.Used/1GB,2)}}

    Write-Section 'Local Shares'
    if ($shares) { $shares | Format-Table -AutoSize } else { Write-Status -Level Warn -Message 'No share data or Get-SmbShare unavailable.' }

    Write-Section 'Mapped/FileSystem Drives'
    $mapped | Format-Table -AutoSize

    $text = @()
    $text += 'Local Shares'
    $text += ($shares | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Mapped/FileSystem Drives'
    $text += ($mapped | Format-Table -AutoSize | Out-String -Width 4096)
    Save-TextReport -Prefix 'Network_SharesAndDrives' -Text ($text -join [Environment]::NewLine) | Out-Null
}

function Get-WifiProfiles {
    $text = netsh wlan show profiles 2>&1 | Out-String
    Show-AndSaveText -Title 'Wi-Fi Profiles' -Prefix 'Network_WifiProfiles' -Text $text
}

function Invoke-Traceroute {
    $target = Read-Host 'Enter host/IP for traceroute'
    if ([string]::IsNullOrWhiteSpace($target)) {
        Write-Status -Level Warn -Message 'Target cannot be empty.'
        return
    }
    $text = tracert $target | Out-String
    Show-AndSaveText -Title "Traceroute: $target" -Prefix 'Network_Traceroute' -Text $text
}

function Invoke-NslookupHelper {
    $name = Read-Host 'Enter DNS name for nslookup'
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Status -Level Warn -Message 'DNS name cannot be empty.'
        return
    }
    $server = Read-Host 'Optional DNS server (leave blank for default)'
    $text = if ([string]::IsNullOrWhiteSpace($server)) {
        nslookup $name | Out-String
    }
    else {
        nslookup $name $server | Out-String
    }
    Show-AndSaveText -Title 'NSLookup Helper' -Prefix 'Network_NSLookup' -Text $text
}

function Get-CommonBackupSources {
    $profile = $env:USERPROFILE
    $roaming = $env:APPDATA
    $local = $env:LOCALAPPDATA

    return @(
        [pscustomobject]@{ Label='Desktop'; Path=(Join-Path $profile 'Desktop'); IsDefault=$true },
        [pscustomobject]@{ Label='Documents'; Path=(Join-Path $profile 'Documents'); IsDefault=$true },
        [pscustomobject]@{ Label='Downloads'; Path=(Join-Path $profile 'Downloads'); IsDefault=$true },
        [pscustomobject]@{ Label='Pictures'; Path=(Join-Path $profile 'Pictures'); IsDefault=$true },
        [pscustomobject]@{ Label='Videos'; Path=(Join-Path $profile 'Videos'); IsDefault=$true },
        [pscustomobject]@{ Label='Music'; Path=(Join-Path $profile 'Music'); IsDefault=$false },
        [pscustomobject]@{ Label='Favorites'; Path=(Join-Path $profile 'Favorites'); IsDefault=$false },
        [pscustomobject]@{ Label='Firefox Profiles'; Path=(Join-Path $roaming 'Mozilla\Firefox\Profiles'); IsDefault=$true },
        [pscustomobject]@{ Label='Chrome User Data'; Path=(Join-Path $local 'Google\Chrome\User Data'); IsDefault=$true },
        [pscustomobject]@{ Label='Edge User Data'; Path=(Join-Path $local 'Microsoft\Edge\User Data'); IsDefault=$true }
    )
}

function Copy-BackupSource {
    param(
        [Parameter(Mandatory)]$Source,
        [Parameter(Mandatory)][string]$DestinationRoot
    )

    if (-not (Test-Path -LiteralPath $Source.Path)) {
        return [pscustomobject]@{
            SourceLabel = $Source.Label
            SourcePath  = $Source.Path
            TargetPath  = $null
            Status      = 'Skipped'
            Details     = 'Source path does not exist.'
        }
    }

    $leafName = Split-Path -Path $Source.Path -Leaf
    if ([string]::IsNullOrWhiteSpace($leafName)) {
        $leafName = Convert-ToSafeFolderName -Name $Source.Label
    }

    $targetName = '{0}_{1}' -f (Convert-ToSafeFolderName -Name $Source.Label), (Convert-ToSafeFolderName -Name $leafName)
    $target = Join-Path -Path $DestinationRoot -ChildPath $targetName

    try {
        if (Test-Path -LiteralPath $target) {
            Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
        }

        if (Test-CommandAvailable -Name robocopy) {
            New-Item -Path $target -ItemType Directory -Force | Out-Null
            $null = robocopy $Source.Path $target /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP
        }
        else {
            Copy-Item -LiteralPath $Source.Path -Destination $target -Recurse -Force -ErrorAction Stop
        }

        return [pscustomobject]@{
            SourceLabel = $Source.Label
            SourcePath  = $Source.Path
            TargetPath  = $target
            Status      = 'Copied'
            Details     = 'Completed'
        }
    }
    catch {
        return [pscustomobject]@{
            SourceLabel = $Source.Label
            SourcePath  = $Source.Path
            TargetPath  = $target
            Status      = 'Error'
            Details     = $_.Exception.Message
        }
    }
}

function Invoke-UserDataBackup {
    Write-Section 'User Data Backup'
    Write-Status -Message 'Build a backup from common personal folders + browser data, then optionally add custom paths.'

    $backupRootInput = Read-Host 'Backup destination root (default: reports folder\User_Backups)'
    $backupRoot = if ([string]::IsNullOrWhiteSpace($backupRootInput)) {
        Join-Path -Path $Script:ReportRoot -ChildPath 'User_Backups'
    }
    else {
        $backupRootInput
    }

    if (-not (Test-Path -LiteralPath $backupRoot)) {
        try {
            New-Item -Path $backupRoot -ItemType Directory -Force | Out-Null
        }
        catch {
            Write-Status -Level Error -Message ("Unable to create backup root '{0}': {1}" -f $backupRoot, $_.Exception.Message)
            return
        }
    }

    $sessionFolder = Join-Path -Path $backupRoot -ChildPath ('{0}_{1}' -f $env:USERNAME, (Get-Timestamp))
    New-Item -Path $sessionFolder -ItemType Directory -Force | Out-Null

    $common = Get-CommonBackupSources
    Write-Host ''
    Write-Host 'Common backup sources:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $common.Count; $i++) {
        $item = $common[$i]
        $defaultText = if ($item.IsDefault) { 'Default' } else { 'Optional' }
        Write-Host ('[{0}] ({1}) {2} -> {3}' -f ($i + 1), $defaultText, $item.Label, $item.Path)
    }
    Write-Host ''
    Write-Host 'Selection options:'
    Write-Host '  - Press Enter = default set'
    Write-Host '  - A = all common locations'
    Write-Host '  - N = none'
    Write-Host '  - Or enter comma-separated numbers (example: 1,2,4,9)'

    $pickRaw = Read-Host 'Select common locations'
    $selectedCommon = @()
    $normalizedPick = if ($null -eq $pickRaw) { '' } else { $pickRaw.Trim().ToUpperInvariant() }
    switch -Regex ($normalizedPick) {
        '^$' { $selectedCommon = $common | Where-Object { $_.IsDefault } }
        '^A$' { $selectedCommon = $common }
        '^N$' { $selectedCommon = @() }
        default {
            $indexes = $pickRaw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ } | Sort-Object -Unique
            foreach ($idx in $indexes) {
                if ($idx -ge 1 -and $idx -le $common.Count) {
                    $selectedCommon += $common[$idx - 1]
                }
            }
        }
    }

    $extraRaw = Read-Host 'Additional file/folder paths (optional, separate with ;)'
    $extraSources = New-Object System.Collections.Generic.List[object]
    if (-not [string]::IsNullOrWhiteSpace($extraRaw)) {
        $extraPaths = $extraRaw -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
        foreach ($path in $extraPaths) {
            $extraSources.Add([pscustomobject]@{
                Label     = 'CustomPath'
                Path      = $path
                IsDefault = $false
            })
        }
    }

    $allSources = New-Object System.Collections.Generic.List[object]
    foreach ($src in $selectedCommon) { $allSources.Add($src) }
    foreach ($src in $extraSources) { $allSources.Add($src) }

    if ($allSources.Count -eq 0) {
        Write-Status -Level Warn -Message 'No backup sources selected.'
        Remove-Item -LiteralPath $sessionFolder -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    Write-Status -Message ("Creating backup in: {0}" -f $sessionFolder)
    $results = foreach ($source in $allSources) {
        Copy-BackupSource -Source $source -DestinationRoot $sessionFolder
    }

    Write-Section 'Backup Results'
    $results | Format-Table -AutoSize

    $summary = @()
    $summary += "Backup root: $backupRoot"
    $summary += "Session folder: $sessionFolder"
    $summary += ('=' * 70)
    $summary += ''
    $summary += ($results | Format-Table -AutoSize | Out-String -Width 4096)

    Save-TextReport -Prefix 'FileSystem_UserBackup' -Text ($summary -join [Environment]::NewLine) | Out-Null
    Write-Status -Level Success -Message 'Backup workflow finished. Review report for copied/skipped/error entries.'
}

function Invoke-RecursiveFileSearch {
    $path = Read-Host 'Search path (default: C:\)'
    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'C:\' }
    $pattern = Read-Host 'File pattern (example: *.log)'
    if ([string]::IsNullOrWhiteSpace($pattern)) { $pattern = '*.*' }

    if (-not (Test-Path -Path $path)) {
        Write-Status -Level Error -Message 'Path does not exist.'
        return
    }

    $data = Get-ChildItem -Path $path -Recurse -File -Filter $pattern -ErrorAction SilentlyContinue |
        Select-Object FullName, Length, LastWriteTime
    Show-AndSaveObject -Title "Recursive File Search ($pattern in $path)" -Prefix 'FileSystem_RecursiveSearch' -Object $data
}

function Invoke-TempCleanup {
    Show-AdminHint -Action 'Temp Cleanup'
    if (-not (Confirm-Action -Prompt 'Delete common temp files now')) {
        Write-Status -Message 'Temp cleanup canceled.'
        return
    }

    $targets = @(
        $env:TEMP,
        (Join-Path $env:WINDIR 'Temp')
    ) | Where-Object { $_ -and (Test-Path $_) }

    $deleted = 0
    foreach ($target in $targets) {
        Write-Status -Message "Cleaning: $target"
        Get-ChildItem -Path $target -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -Recurse -ErrorAction Stop
                $deleted++
            }
            catch {
            }
        }
    }

    $text = "Temp cleanup completed. Deleted items: $deleted"
    Show-AndSaveText -Title 'Temp Cleanup Result' -Prefix 'FileSystem_TempCleanup' -Text $text
}

function Find-LargeFiles {
    $path = Read-Host 'Path to scan for large files (default: C:\)'
    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'C:\' }
    $minMbRaw = Read-Host 'Minimum file size in MB (default: 100)'
    $minMb = 100
    if ($minMbRaw -match '^[0-9]+$') { $minMb = [int]$minMbRaw }

    if (-not (Test-Path -Path $path)) {
        Write-Status -Level Error -Message 'Path does not exist.'
        return
    }

    $data = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Length -ge ($minMb * 1MB) } |
        Sort-Object Length -Descending |
        Select-Object -First 200 FullName,
            @{Name='SizeMB';Expression={[math]::Round($_.Length/1MB,2)}},
            LastWriteTime

    Show-AndSaveObject -Title "Large Files Finder (>$minMb MB in $path)" -Prefix 'FileSystem_LargeFiles' -Object $data
}

function Get-RecentFilesListing {
    $path = Read-Host 'Path to scan recent files (default: C:\Users)'
    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'C:\Users' }

    if (-not (Test-Path -Path $path)) {
        Write-Status -Level Error -Message 'Path does not exist.'
        return
    }

    $data = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 200 FullName, LastWriteTime,
            @{Name='SizeKB';Expression={[math]::Round($_.Length/1KB,2)}}

    Show-AndSaveObject -Title "Recent Files Listing ($path)" -Prefix 'FileSystem_RecentFiles' -Object $data
}

function Export-DirectoryTree {
    $path = Read-Host 'Path to export tree from (default: C:\)'
    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'C:\' }

    if (-not (Test-Path -Path $path)) {
        Write-Status -Level Error -Message 'Path does not exist.'
        return
    }

    $export = New-ReportPath -Prefix 'FileSystem_DirectoryTree'
    cmd /c tree "$path" /f /a > "$export"
    $Script:LastReports.Add($export)
    Write-Status -Level Success -Message "Directory tree exported: $export"
}

function Get-DriveFreeSpaceSummary {
    $data = Get-PSDrive -PSProvider FileSystem |
        Select-Object Name, Root,
            @{Name='UsedGB';Expression={[math]::Round($_.Used/1GB,2)}},
            @{Name='FreeGB';Expression={[math]::Round($_.Free/1GB,2)}}
    Show-AndSaveObject -Title 'Drive Free Space Summary' -Prefix 'FileSystem_DriveSpace' -Object $data
}

function Find-DuplicateFileNameCandidates {
    $path = Read-Host 'Path for duplicate filename scan (default: C:\Users)'
    if ([string]::IsNullOrWhiteSpace($path)) { $path = 'C:\Users' }

    if (-not (Test-Path -Path $path)) {
        Write-Status -Level Error -Message 'Path does not exist.'
        return
    }

    $dupes = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
        Group-Object Name |
        Where-Object { $_.Count -gt 1 } |
        Sort-Object Count -Descending |
        Select-Object -First 200 Name, Count

    Show-AndSaveObject -Title 'Duplicate Filename Candidates' -Prefix 'FileSystem_DuplicateNameCandidates' -Object $dupes
}

function Invoke-PathReport {
    $path = Read-Host 'Enter path for quick report'
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -Path $path)) {
        Write-Status -Level Error -Message 'Valid path is required.'
        return
    }

    $items = Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue
    $fileCount = ($items | Where-Object { -not $_.PSIsContainer }).Count
    $dirCount = ($items | Where-Object { $_.PSIsContainer }).Count
    $size = ($items | Where-Object { -not $_.PSIsContainer } | Measure-Object Length -Sum).Sum

    $text = @(
        "Path: $path",
        "Directories: $dirCount",
        "Files: $fileCount",
        ('Total Size GB: {0:N2}' -f ($size / 1GB))
    ) -join [Environment]::NewLine

    Show-AndSaveText -Title 'User-Selected Path Report' -Prefix 'FileSystem_PathReport' -Text $text
}

function Get-LocalUsersList {
    if (Test-CommandAvailable -Name Get-LocalUser) {
        $data = Get-LocalUser -ErrorAction SilentlyContinue | Select-Object Name, Enabled, LastLogon, PasswordRequired
        Show-AndSaveObject -Title 'Local Users' -Prefix 'Users_LocalUsers' -Object $data
    }
    else {
        $text = net user | Out-String
        Show-AndSaveText -Title 'Local Users' -Prefix 'Users_LocalUsers' -Text $text
    }
}

function Get-AdministratorsGroupMembers {
    if (Test-CommandAvailable -Name Get-LocalGroupMember) {
        $data = Get-LocalGroupMember -Group 'Administrators' -ErrorAction SilentlyContinue |
            Select-Object Name, ObjectClass, PrincipalSource
        Show-AndSaveObject -Title 'Administrators Group Members' -Prefix 'Users_AdminGroupMembers' -Object $data
    }
    else {
        $text = net localgroup administrators | Out-String
        Show-AndSaveText -Title 'Administrators Group Members' -Prefix 'Users_AdminGroupMembers' -Text $text
    }
}

function Get-LastLogonInfo {
    $profiles = Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue |
        Where-Object { $_.LocalPath -like 'C:\Users\*' } |
        Select-Object LocalPath, LastUseTime, Loaded

    Show-AndSaveObject -Title 'Last Logon Info (Profile Last Use Time)' -Prefix 'Users_LastLogonInfo' -Object $profiles
}

function Get-ProfileFolderListing {
    $path = 'C:\Users'
    if (-not (Test-Path -Path $path)) {
        Write-Status -Level Warn -Message 'C:\Users path unavailable.'
        return
    }

    $data = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue |
        Select-Object Name, FullName, LastWriteTime
    Show-AndSaveObject -Title 'Profile Folder Listing' -Prefix 'Users_ProfileFolders' -Object $data
}

function Invoke-QuickUserAudit {
    $localUsers = if (Test-CommandAvailable -Name Get-LocalUser) {
        Get-LocalUser -ErrorAction SilentlyContinue | Select-Object Name, Enabled, LastLogon
    }
    else {
        $null
    }

    $admins = if (Test-CommandAvailable -Name Get-LocalGroupMember) {
        Get-LocalGroupMember -Group 'Administrators' -ErrorAction SilentlyContinue | Select-Object Name, ObjectClass
    }
    else {
        $null
    }

    $profiles = Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue |
        Where-Object { $_.LocalPath -like 'C:\Users\*' } |
        Select-Object LocalPath, LastUseTime, Loaded

    $text = @()
    $text += 'Quick User/Account Audit'
    $text += ('=' * 40)
    $text += ''
    $text += 'Local Users:'
    $text += ($localUsers | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Administrators Group:'
    $text += ($admins | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Profile Last Use:'
    $text += ($profiles | Format-Table -AutoSize | Out-String -Width 4096)

    Show-AndSaveText -Title 'Quick User/Account Audit Report' -Prefix 'Users_QuickAudit' -Text ($text -join [Environment]::NewLine)
}

function Get-BitLockerStatus {
    if (Test-CommandAvailable -Name Get-BitLockerVolume) {
        $data = Get-BitLockerVolume -ErrorAction SilentlyContinue |
            Select-Object MountPoint, VolumeType, VolumeStatus, ProtectionStatus, EncryptionPercentage
        Show-AndSaveObject -Title 'BitLocker Status' -Prefix 'Security_BitLocker' -Object $data
    }
    else {
        $text = manage-bde -status | Out-String
        Show-AndSaveText -Title 'BitLocker Status' -Prefix 'Security_BitLocker' -Text $text
    }
}

function Get-FirewallStatus {
    if (Test-CommandAvailable -Name Get-NetFirewallProfile) {
        $data = Get-NetFirewallProfile -ErrorAction SilentlyContinue |
            Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction
        Show-AndSaveObject -Title 'Firewall Profile Status' -Prefix 'Security_Firewall' -Object $data
    }
    else {
        $text = netsh advfirewall show allprofiles | Out-String
        Show-AndSaveText -Title 'Firewall Profile Status' -Prefix 'Security_Firewall' -Text $text
    }
}

function Get-DefenderStatus {
    if (Test-CommandAvailable -Name Get-MpComputerStatus) {
        $data = Get-MpComputerStatus -ErrorAction SilentlyContinue |
            Select-Object AMServiceEnabled, AntivirusEnabled, RealTimeProtectionEnabled,
                NISEnabled, AntivirusSignatureLastUpdated, QuickScanStartTime
        Show-AndSaveObject -Title 'Defender Status' -Prefix 'Security_Defender' -Object $data -Format List
    }
    else {
        Write-Status -Level Warn -Message 'Defender cmdlets unavailable on this system.'
    }
}

function Get-SuspiciousPersistenceCheck {
    $startup = Get-CimInstance Win32_StartupCommand -ErrorAction SilentlyContinue |
        Select-Object Name, Command, Location, User

    $tasks = if (Test-CommandAvailable -Name Get-ScheduledTask) {
        Get-ScheduledTask -ErrorAction SilentlyContinue |
            Where-Object {
                ($_.Triggers | Where-Object { $_.CimClass.CimClassName -match 'Logon|Boot' })
            } |
            Select-Object TaskName, TaskPath, State
    }

    Write-Section 'Startup/Persistence Check - Startup Commands'
    if ($startup) { $startup | Format-Table -AutoSize } else { Write-Status -Level Warn -Message 'No startup command data found.' }

    Write-Section 'Startup/Persistence Check - Logon/Boot Scheduled Tasks'
    if ($tasks) { $tasks | Format-Table -AutoSize } else { Write-Status -Level Warn -Message 'No suspicious scheduled task patterns found.' }

    $text = @()
    $text += 'Startup Commands'
    $text += ($startup | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Boot/Logon Tasks'
    $text += ($tasks | Format-Table -AutoSize | Out-String -Width 4096)
    Save-TextReport -Prefix 'Security_PersistenceCheck' -Text ($text -join [Environment]::NewLine) | Out-Null
}

function Invoke-SecuritySnapshot {
    Write-Section 'Quick Security Posture Snapshot'
    $bit = if (Test-CommandAvailable -Name Get-BitLockerVolume) {
        Get-BitLockerVolume -ErrorAction SilentlyContinue | Select-Object MountPoint, ProtectionStatus, EncryptionPercentage
    }

    $fw = if (Test-CommandAvailable -Name Get-NetFirewallProfile) {
        Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select-Object Name, Enabled
    }

    $def = if (Test-CommandAvailable -Name Get-MpComputerStatus) {
        Get-MpComputerStatus -ErrorAction SilentlyContinue | Select-Object AntivirusEnabled, RealTimeProtectionEnabled, AntivirusSignatureLastUpdated
    }

    $text = @()
    $text += 'Quick Security Posture Snapshot'
    $text += ('=' * 40)
    $text += ''
    $text += 'BitLocker:'
    $text += ($bit | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Firewall Profiles:'
    $text += ($fw | Format-Table -AutoSize | Out-String -Width 4096)
    $text += ''
    $text += 'Defender:'
    $text += ($def | Format-Table -AutoSize | Out-String -Width 4096)

    Show-AndSaveText -Title 'Quick Security Posture Snapshot' -Prefix 'Security_Snapshot' -Text ($text -join [Environment]::NewLine)
}

function Get-ServicesList {
    $data = Get-Service -ErrorAction SilentlyContinue | Sort-Object DisplayName | Select-Object Name, DisplayName, Status, StartType
    Show-AndSaveObject -Title 'Services List' -Prefix 'Services_All' -Object $data
}

function Get-ServicesByState {
    $state = Read-Host 'Enter state filter (Running/Stopped)'
    if ($state -notin @('Running','Stopped')) {
        Write-Status -Level Warn -Message 'Use Running or Stopped.'
        return
    }

    $data = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq $state } |
        Select-Object Name, DisplayName, Status, StartType

    Show-AndSaveObject -Title "Services Filtered by $state" -Prefix "Services_$state" -Object $data
}

function Restart-ServiceByName {
    Show-AdminHint -Action 'Restart Service'
    $name = Read-Host 'Enter service name to restart'
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Status -Level Warn -Message 'Service name is required.'
        return
    }

    if (-not (Confirm-Action -Prompt "Restart service '$name'")) {
        Write-Status -Message 'Service restart canceled.'
        return
    }

    Invoke-Safe -ActionName 'Restart-Service' -ScriptBlock {
        Restart-Service -Name $name -Force -ErrorAction Stop
        Write-Status -Level Success -Message "Service restarted: $name"
        Save-TextReport -Prefix 'Services_Restart' -Text "Restarted service: $name" | Out-Null
    }
}

function Get-ServiceDetails {
    $name = Read-Host 'Enter service name'
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Status -Level Warn -Message 'Service name is required.'
        return
    }

    $svc = Get-Service -Name $name -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, Status, StartType
    if (-not $svc) {
        Write-Status -Level Warn -Message "Service not found: $name"
        return
    }

    Show-AndSaveObject -Title "Service Details: $name" -Prefix 'Services_Details' -Object $svc -Format List
}

function Invoke-CommonServiceQuickActions {
    Show-Header
    Write-Section 'Common Service Quick Actions'
    Write-Host '1. Restart Spooler'
    Write-Host '2. Restart wuauserv'
    Write-Host '3. Restart BITS'
    Write-Host '4. Query Spooler'
    Write-Host '5. Query wuauserv'
    Write-Host '0. Back'

    $choice = Read-Host 'Select option'
    switch ($choice.ToUpperInvariant()) {
        '1' {
            Show-AdminHint -Action 'Restart Spooler'
            if (Confirm-Action -Prompt 'Restart Spooler service') {
                Invoke-Safe -ActionName 'Restart Spooler' -ScriptBlock { Restart-Service -Name spooler -Force -ErrorAction Stop }
                Save-TextReport -Prefix 'Services_QuickAction' -Text 'Restarted service: spooler' | Out-Null
            }
        }
        '2' {
            Show-AdminHint -Action 'Restart wuauserv'
            if (Confirm-Action -Prompt 'Restart wuauserv service') {
                Invoke-Safe -ActionName 'Restart wuauserv' -ScriptBlock { Restart-Service -Name wuauserv -Force -ErrorAction Stop }
                Save-TextReport -Prefix 'Services_QuickAction' -Text 'Restarted service: wuauserv' | Out-Null
            }
        }
        '3' {
            Show-AdminHint -Action 'Restart BITS'
            if (Confirm-Action -Prompt 'Restart BITS service') {
                Invoke-Safe -ActionName 'Restart BITS' -ScriptBlock { Restart-Service -Name bits -Force -ErrorAction Stop }
                Save-TextReport -Prefix 'Services_QuickAction' -Text 'Restarted service: BITS' | Out-Null
            }
        }
        '4' {
            Invoke-Safe -ActionName 'Query Spooler' -ScriptBlock { Get-Service -Name spooler | Format-List * }
        }
        '5' {
            Invoke-Safe -ActionName 'Query wuauserv' -ScriptBlock { Get-Service -Name wuauserv | Format-List * }
        }
        default { }
    }
}
function Invoke-FullTriageReport {
    Write-Section 'Running Full Triage Report'
    Write-Status -Message 'Collecting representative data from all categories.'

    Invoke-Safe -ActionName 'System Audit' -ScriptBlock { Invoke-FullSystemAudit }
    Invoke-Safe -ActionName 'IP Config' -ScriptBlock { Get-IPConfiguration }
    Invoke-Safe -ActionName 'Drive Space' -ScriptBlock { Get-DriveFreeSpaceSummary }
    Invoke-Safe -ActionName 'User Audit' -ScriptBlock { Invoke-QuickUserAudit }
    Invoke-Safe -ActionName 'Security Snapshot' -ScriptBlock { Invoke-SecuritySnapshot }
    Invoke-Safe -ActionName 'Services List' -ScriptBlock { Get-ServicesList }

    Write-Status -Level Success -Message 'Full triage report run complete.'
}

function Invoke-CategoryOnlyReport {
    Show-Header
    Write-Section 'Category-Only Report Runner'
    Write-Host '1. System'
    Write-Host '2. Network'
    Write-Host '3. File System'
    Write-Host '4. Users'
    Write-Host '5. Security'
    Write-Host '6. Services'
    Write-Host '0. Back'

    $choice = Read-Host 'Select category'
    switch ($choice.ToUpperInvariant()) {
        '1' { Invoke-FullSystemAudit }
        '2' {
            Get-IPConfiguration
            Get-AdapterSummary
            Get-RouteTable
            Get-ActiveTcpConnections
        }
        '3' {
            Get-DriveFreeSpaceSummary
            Get-RecentFilesListing
        }
        '4' { Invoke-QuickUserAudit }
        '5' { Invoke-SecuritySnapshot }
        '6' {
            Get-ServicesList
            Get-ServicesByState
        }
        default { }
    }
}

function Export-CombinedSummaryReport {
    $target = Read-Host 'Export location (blank = reports folder)'
    if ([string]::IsNullOrWhiteSpace($target)) { $target = $Script:ReportRoot }

    if (-not (Test-Path -Path $target)) {
        try {
            New-Item -Path $target -ItemType Directory -Force | Out-Null
        }
        catch {
            Write-Status -Level Error -Message 'Unable to create export location.'
            return
        }
    }

    $outFile = Join-Path -Path $target -ChildPath ("CombinedSummary_{0}.txt" -f (Get-Timestamp))
    $recent = Get-ChildItem -Path $Script:ReportRoot -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 25

    $body = @()
    $body += "Combined Summary - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $body += ('=' * 70)
    foreach ($file in $recent) {
        $body += ''
        $body += ("### File: {0}" -f $file.FullName)
        $body += ('-' * 70)
        $body += (Get-Content -Path $file.FullName -ErrorAction SilentlyContinue)
    }

    $body | Set-Content -Path $outFile -Encoding UTF8
    $Script:LastReports.Add($outFile)
    Write-Status -Level Success -Message "Combined summary exported: $outFile"
}

function Open-ReportsFolder {
    if (Test-Path -Path $Script:ReportRoot) {
        Invoke-Item -Path $Script:ReportRoot
    }
}

function Show-LastReports {
    $reports = Get-ChildItem -Path $Script:ReportRoot -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 20 Name, LastWriteTime, Length, FullName

    Write-Section 'Last Generated Reports'
    if ($reports) {
        $reports | Format-Table -AutoSize
    }
    else {
        Write-Status -Level Warn -Message 'No reports found.'
    }
}

function New-TechnicianNote {
    $export = Read-Host 'Optional note export location (blank = reports folder)'
    if ([string]::IsNullOrWhiteSpace($export)) { $export = $Script:ReportRoot }
    if (-not (Test-Path -Path $export)) {
        New-Item -Path $export -ItemType Directory -Force | Out-Null
    }

    $outFile = Join-Path -Path $export -ChildPath ("TechnicianNote_{0}.txt" -f (Get-Timestamp))
    Write-Host 'Enter note lines. Type a single period (.) on a line to finish.' -ForegroundColor Gray
    $lines = New-Object System.Collections.Generic.List[string]

    while ($true) {
        $line = Read-Host ''
        if ($line -eq '.') { break }
        $lines.Add($line)
    }

    $content = @(
        "Technician Note - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "Host: $env:COMPUTERNAME",
        "User: $env:USERNAME",
        ('=' * 50),
        ''
    ) + $lines

    $content | Set-Content -Path $outFile -Encoding UTF8
    $Script:LastReports.Add($outFile)
    Write-Status -Level Success -Message "Technician note saved: $outFile"
}

function Stop-ProcessByPID {
    Show-AdminHint -Action 'Kill Process by PID'
    $pidInput = Read-Host 'Enter process PID'
    if ($pidInput -notmatch '^[0-9]+$') {
        Write-Status -Level Warn -Message 'PID must be numeric.'
        return
    }

    $pidValue = [int]$pidInput
    $proc = Get-Process -Id $pidValue -ErrorAction SilentlyContinue
    if (-not $proc) {
        Write-Status -Level Warn -Message "Process not found: PID $pidValue"
        return
    }

    Write-Host ("Target: {0} (PID {1})" -f $proc.ProcessName, $proc.Id)
    if (-not (Confirm-Action -Prompt 'Terminate this process')) {
        Write-Status -Message 'Process termination canceled.'
        return
    }

    Invoke-Safe -ActionName 'Stop-Process' -ScriptBlock {
        Stop-Process -Id $pidValue -Force -ErrorAction Stop
        Write-Status -Level Success -Message "Process terminated: PID $pidValue"
        Save-TextReport -Prefix 'Tools_ProcessKill' -Text "Killed process PID: $pidValue" | Out-Null
    }
}

function Export-DriversWithDism {
    Show-AdminHint -Action 'Driver Export (DISM)'
    if (-not (Test-CommandAvailable -Name dism)) {
        Write-Status -Level Warn -Message 'DISM is unavailable on this system.'
        return
    }

    $dest = Read-Host 'Driver export destination (blank = reports folder\DriverExport_<timestamp>)'
    if ([string]::IsNullOrWhiteSpace($dest)) {
        $dest = Join-Path -Path $Script:ReportRoot -ChildPath ("DriverExport_{0}" -f (Get-Timestamp))
    }

    if (-not (Test-Path -Path $dest)) {
        New-Item -Path $dest -ItemType Directory -Force | Out-Null
    }

    $output = dism /online /export-driver /destination:"$dest" 2>&1 | Out-String
    Show-AndSaveText -Title 'Driver Export (DISM)' -Prefix 'Tools_DriverExport' -Text ($output + [Environment]::NewLine + "Destination: $dest")
}

function Get-HostnameSerialSummary {
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $text = @(
        "Hostname: $env:COMPUTERNAME",
        "Manufacturer: $($cs.Manufacturer)",
        "Model: $($cs.Model)",
        "Serial: $($bios.SerialNumber)"
    ) -join [Environment]::NewLine

    Show-AndSaveText -Title 'Hostname and Serial Summary' -Prefix 'Tools_HostnameSerial' -Text $text
}

function Copy-TextToClipboardHelper {
    if (-not (Test-CommandAvailable -Name Set-Clipboard)) {
        Write-Status -Level Warn -Message 'Set-Clipboard is unavailable on this system.'
        return
    }

    $recent = Get-ChildItem -Path $Script:ReportRoot -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 10
    if (-not $recent) {
        Write-Status -Level Warn -Message 'No reports available to copy.'
        return
    }

    Write-Section 'Select Report to Copy to Clipboard'
    for ($i = 0; $i -lt $recent.Count; $i++) {
        Write-Host ('{0}. {1}' -f ($i + 1), $recent[$i].Name)
    }

    $pick = Read-Host 'Enter number'
    if ($pick -notmatch '^[0-9]+$') {
        Write-Status -Level Warn -Message 'Invalid selection.'
        return
    }

    $idx = [int]$pick - 1
    if ($idx -lt 0 -or $idx -ge $recent.Count) {
        Write-Status -Level Warn -Message 'Selection out of range.'
        return
    }

    $content = Get-Content -Path $recent[$idx].FullName -Raw -ErrorAction SilentlyContinue
    Set-Clipboard -Value $content
    Write-Status -Level Success -Message ("Copied report to clipboard: {0}" -f $recent[$idx].Name)
}

function Invoke-BuiltInToolsMenu {
    Show-Header
    Write-Section 'Built-in Windows Tools'
    Write-Host '1. msinfo32'
    Write-Host '2. eventvwr'
    Write-Host '3. devmgmt.msc'
    Write-Host '4. services.msc'
    Write-Host '5. ncpa.cpl'
    Write-Host '6. compmgmt.msc'
    Write-Host '0. Back'

    $choice = Read-Host 'Select tool'
    switch ($choice.ToUpperInvariant()) {
        '1' { Start-Process msinfo32 }
        '2' { Start-Process eventvwr }
        '3' { Start-Process devmgmt.msc }
        '4' { Start-Process services.msc }
        '5' { Start-Process ncpa.cpl }
        '6' { Start-Process compmgmt.msc }
        default { }
    }
}

$Script:MenuActions = [ordered]@{
    'System' = @(
        @{ Label='Full system audit'; Handler='Invoke-FullSystemAudit' },
        @{ Label='Basic hardware info'; Handler='Get-BasicHardwareInfo' },
        @{ Label='OS info'; Handler='Get-OSInfo' },
        @{ Label='CPU info'; Handler='Get-CPUInfo' },
        @{ Label='RAM info'; Handler='Get-RAMInfo' },
        @{ Label='Disk info'; Handler='Get-DiskInfo' },
        @{ Label='Top processes by CPU/RAM'; Handler='Get-TopProcesses' },
        @{ Label='Installed software inventory'; Handler='Get-InstalledSoftwareInventory' },
        @{ Label='Installed Windows updates/hotfixes'; Handler='Get-HotfixInventory' },
        @{ Label='Event log errors from last 24 hours'; Handler='Get-EventLogErrors24h' },
        @{ Label='Startup entries audit'; Handler='Get-StartupEntriesAudit' },
        @{ Label='Scheduled tasks summary'; Handler='Get-ScheduledTasksSummary' },
        @{ Label='Environment info'; Handler='Get-EnvironmentInfo' },
        @{ Label='System uptime'; Handler='Get-SystemUptime' }
    )
    'Network' = @(
        @{ Label='IP configuration'; Handler='Get-IPConfiguration' },
        @{ Label='Adapter summary'; Handler='Get-AdapterSummary' },
        @{ Label='ARP table'; Handler='Get-ArpTable' },
        @{ Label='Neighbor table'; Handler='Get-NeighborTable' },
        @{ Label='Route table'; Handler='Get-RouteTable' },
        @{ Label='DNS resolution test'; Handler='Invoke-DnsResolutionTest' },
        @{ Label='Ping test'; Handler='Invoke-PingTest' },
        @{ Label='Open/listening ports'; Handler='Get-ListeningPorts' },
        @{ Label='Active TCP connections'; Handler='Get-ActiveTcpConnections' },
        @{ Label='Network reset option'; Handler='Invoke-NetworkReset' },
        @{ Label='Local shares and mapped drives'; Handler='Get-LocalSharesAndMappedDrives' },
        @{ Label='Wi-Fi profile list'; Handler='Get-WifiProfiles' },
        @{ Label='Traceroute option'; Handler='Invoke-Traceroute' },
        @{ Label='nslookup helper'; Handler='Invoke-NslookupHelper' }
    )
    'File System' = @(
        @{ Label='User data backup (common + custom paths)'; Handler='Invoke-UserDataBackup' },
        @{ Label='Recursive file search by pattern'; Handler='Invoke-RecursiveFileSearch' },
        @{ Label='Temp cleanup'; Handler='Invoke-TempCleanup' },
        @{ Label='Large files finder'; Handler='Find-LargeFiles' },
        @{ Label='Recent files listing'; Handler='Get-RecentFilesListing' },
        @{ Label='Export directory tree to text'; Handler='Export-DirectoryTree' },
        @{ Label='Drive free space summary'; Handler='Get-DriveFreeSpaceSummary' },
        @{ Label='Duplicate candidate finder (filename only)'; Handler='Find-DuplicateFileNameCandidates' },
        @{ Label='User-selected path report'; Handler='Invoke-PathReport' }
    )
    'Users' = @(
        @{ Label='Local users'; Handler='Get-LocalUsersList' },
        @{ Label='Administrators group members'; Handler='Get-AdministratorsGroupMembers' },
        @{ Label='Last logon info where available'; Handler='Get-LastLogonInfo' },
        @{ Label='Profile folder listing'; Handler='Get-ProfileFolderListing' },
        @{ Label='Quick user/account audit report'; Handler='Invoke-QuickUserAudit' }
    )
    'Security' = @(
        @{ Label='BitLocker status'; Handler='Get-BitLockerStatus' },
        @{ Label='Firewall profile status'; Handler='Get-FirewallStatus' },
        @{ Label='Defender status'; Handler='Get-DefenderStatus' },
        @{ Label='Suspicious startup/persistence checks'; Handler='Get-SuspiciousPersistenceCheck' },
        @{ Label='Quick security posture snapshot'; Handler='Invoke-SecuritySnapshot' }
    )
    'Services' = @(
        @{ Label='List services'; Handler='Get-ServicesList' },
        @{ Label='Filter services by running/stopped'; Handler='Get-ServicesByState' },
        @{ Label='Restart selected service by name'; Handler='Restart-ServiceByName' },
        @{ Label='Query service details'; Handler='Get-ServiceDetails' },
        @{ Label='Common service quick actions (spooler/wuauserv/BITS)'; Handler='Invoke-CommonServiceQuickActions' }
    )
    'Reports' = @(
        @{ Label='Run full triage report'; Handler='Invoke-FullTriageReport' },
        @{ Label='Run category-only reports'; Handler='Invoke-CategoryOnlyReport' },
        @{ Label='Export combined summary report'; Handler='Export-CombinedSummaryReport' },
        @{ Label='Open reports folder'; Handler='Open-ReportsFolder' },
        @{ Label='Display last generated reports'; Handler='Show-LastReports' },
        @{ Label='Create technician note file'; Handler='New-TechnicianNote' }
    )
    'Tools/Utilities' = @(
        @{ Label='Process kill by PID'; Handler='Stop-ProcessByPID' },
        @{ Label='Driver export using DISM'; Handler='Export-DriversWithDism' },
        @{ Label='Hostname and serial summary'; Handler='Get-HostnameSerialSummary' },
        @{ Label='Clipboard-safe plain-text output helper'; Handler='Copy-TextToClipboardHelper' },
        @{ Label='Launch built-in Windows tools'; Handler='Invoke-BuiltInToolsMenu' }
    )
}

function Get-AllActions {
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($category in $Script:MenuActions.Keys) {
        foreach ($item in $Script:MenuActions[$category]) {
            $all.Add([pscustomobject]@{
                Category = $category
                Label    = $item.Label
                Handler  = $item.Handler
            })
        }
    }
    return $all
}

function Invoke-ActionHandler {
    param(
        [Parameter(Mandatory)][string]$Handler,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not (Get-Command -Name $Handler -CommandType Function -ErrorAction SilentlyContinue)) {
        Write-Status -Level Error -Message "Action unavailable: $Label ($Handler)"
        return
    }

    Invoke-Safe -ActionName $Label -ScriptBlock { & $Handler }
}

function Show-CategoryMenu {
    param([Parameter(Mandatory)][string]$Category)

    while ($true) {
        Show-Header
        Write-Section "$Category Menu"

        $actions = if ($Category -eq 'All') { Get-AllActions } else {
            $Script:MenuActions[$Category] | ForEach-Object {
                [pscustomobject]@{
                    Category = $Category
                    Label    = $_.Label
                    Handler  = $_.Handler
                }
            }
        }

        for ($i = 0; $i -lt $actions.Count; $i++) {
            if ($Category -eq 'All') {
                Write-Host ('{0}. [{1}] {2}' -f ($i + 1), $actions[$i].Category, $actions[$i].Label)
            }
            else {
                Write-Host ('{0}. {1}' -f ($i + 1), $actions[$i].Label)
            }
        }

        Write-Host '0. Back'
        $choice = Read-Host 'Select action'

        if ($choice -eq '0') {
            return
        }

        if ($choice -notmatch '^[0-9]+$') {
            Write-Status -Level Warn -Message 'Invalid selection.'
            Pause-Toolkit
            continue
        }

        $index = [int]$choice - 1
        if ($index -lt 0 -or $index -ge $actions.Count) {
            Write-Status -Level Warn -Message 'Selection out of range.'
            Pause-Toolkit
            continue
        }

        $selected = $actions[$index]
        Show-Header
        Write-Section ("Executing: {0}" -f $selected.Label)
        Invoke-ActionHandler -Handler $selected.Handler -Label $selected.Label
        Pause-Toolkit
    }
}

function Show-MainMenu {
    while ($true) {
        Show-Header
        Write-Section 'Main Category Menu'
        Write-Host '1. All'
        Write-Host '2. System'
        Write-Host '3. Network'
        Write-Host '4. File System'
        Write-Host '5. Users'
        Write-Host '6. Security'
        Write-Host '7. Services'
        Write-Host '8. Reports'
        Write-Host '9. Tools/Utilities'
        Write-Host 'X. Exit'

        $choice = Read-Host 'Select category'
        switch ($choice.ToUpperInvariant()) {
            '1' { Show-CategoryMenu -Category 'All' }
            '2' { Show-CategoryMenu -Category 'System' }
            '3' { Show-CategoryMenu -Category 'Network' }
            '4' { Show-CategoryMenu -Category 'File System' }
            '5' { Show-CategoryMenu -Category 'Users' }
            '6' { Show-CategoryMenu -Category 'Security' }
            '7' { Show-CategoryMenu -Category 'Services' }
            '8' { Show-CategoryMenu -Category 'Reports' }
            '9' { Show-CategoryMenu -Category 'Tools/Utilities' }
            'X' { return }
            default {
                Write-Status -Level Warn -Message 'Unknown selection.'
                Pause-Toolkit
            }
        }
    }
}

Initialize-ReportSession
Show-MainMenu
