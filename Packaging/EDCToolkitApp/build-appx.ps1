[CmdletBinding()]
param(
    [string]$Version = '0.1.0.0',
    [string]$Publisher = 'CN=EDCToolkitDev',
    [string]$PackageName = 'EDCToolkit.Beta',
    [switch]$SkipSign
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ToolPath {
    param(
        [Parameter(Mandatory)][string]$ToolName,
        [Parameter(Mandatory)][string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if (Test-Path -Path $candidate) { return $candidate }
    }
    throw "Required tool not found: $ToolName"
}

function Ensure-Directory {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$packageRoot = $PSScriptRoot
$buildRoot = Join-Path $packageRoot 'build'
$outputRoot = Join-Path $packageRoot 'output'
$certRoot = Join-Path $packageRoot 'certs'
$layoutRoot = Join-Path $buildRoot 'layout'
$assetsRoot = Join-Path $layoutRoot 'Assets'
$scriptsRoot = Join-Path $layoutRoot 'Scripts'

Ensure-Directory -Path $buildRoot
Ensure-Directory -Path $outputRoot
Ensure-Directory -Path $certRoot
if (Test-Path -Path $layoutRoot) {
    Remove-Item -Path $layoutRoot -Recurse -Force
}
Ensure-Directory -Path $layoutRoot
Ensure-Directory -Path $assetsRoot
Ensure-Directory -Path $scriptsRoot

$csc = Get-ToolPath -ToolName 'csc.exe' -Candidates @(
    'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe',
    'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe'
)
$makeAppx = Get-ToolPath -ToolName 'makeappx.exe' -Candidates @(
    'C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\makeappx.exe',
    'C:\Program Files (x86)\Windows Kits\10\App Certification Kit\makeappx.exe'
)
$signTool = Get-ToolPath -ToolName 'signtool.exe' -Candidates @(
    'C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe',
    'C:\Program Files (x86)\Windows Kits\10\App Certification Kit\signtool.exe'
)

$launcherSource = Join-Path $packageRoot 'Launcher\EDCToolkitLauncher.cs'
$launcherExe = Join-Path $layoutRoot 'EDCToolkitLauncher.exe'
& $csc /nologo /target:winexe /out:$launcherExe $launcherSource

Copy-Item -Path (Join-Path $repoRoot 'Scripts\EDCtoolkit') -Destination $scriptsRoot -Recurse -Force
if (Test-Path -Path (Join-Path $scriptsRoot 'EDCtoolkit\EDC_Reports')) {
    Remove-Item -Path (Join-Path $scriptsRoot 'EDCtoolkit\EDC_Reports') -Recurse -Force
}

$logoSource = Join-Path $repoRoot 'Scripts\EDCtoolkit\logo-n.png'
$assetMap = @(
    'StoreLogo.png',
    'Square44x44Logo.png',
    'Square150x150Logo.png',
    'Wide310x150Logo.png',
    'Square310x310Logo.png'
)
foreach ($asset in $assetMap) {
    Copy-Item -Path $logoSource -Destination (Join-Path $assetsRoot $asset) -Force
}

$manifestTemplate = Get-Content -Raw -Encoding UTF8 (Join-Path $packageRoot 'AppxManifest.xml')
$manifest = $manifestTemplate.Replace('Version="0.1.0.0"', ('Version="{0}"' -f $Version)).Replace('Publisher="CN=EDCToolkitDev"', ('Publisher="{0}"' -f $Publisher)).Replace('Name="EDCToolkit.Beta"', ('Name="{0}"' -f $PackageName))
Set-Content -Path (Join-Path $layoutRoot 'AppxManifest.xml') -Value $manifest -Encoding UTF8

$appxPath = Join-Path $outputRoot ('EDCToolkit_{0}.appx' -f $Version)
if (Test-Path -Path $appxPath) {
    Remove-Item -Path $appxPath -Force
}

& $makeAppx pack /d $layoutRoot /p $appxPath /o

if (-not $SkipSign) {
    try {
        $certName = 'EDCToolkitDev'
        $certPassword = ConvertTo-SecureString -String 'EDCToolkitDev123!' -AsPlainText -Force
        $pfxPath = Join-Path $certRoot 'EDCToolkitDev.pfx'
        $cerPath = Join-Path $certRoot 'EDCToolkitDev.cer'

        if (-not (Test-Path -Path $pfxPath)) {
            $cert = New-SelfSignedCertificate -Subject $Publisher -FriendlyName $certName -CertStoreLocation 'Cert:\CurrentUser\My' -KeyExportPolicy Exportable -HashAlgorithm 'SHA256' -KeyAlgorithm RSA -KeyLength 2048 -NotAfter (Get-Date).AddYears(3)
            Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $certPassword | Out-Null
            Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
        }

        & $signTool sign /fd SHA256 /f $pfxPath /p 'EDCToolkitDev123!' $appxPath
    }
    catch {
        Write-Warning ("Package built but signing was skipped: {0}" -f $_.Exception.Message)
        Write-Warning "Run with -SkipSign or provide your own signing certificate for installable test builds."
    }
}

Write-Host "Built package: $appxPath"
