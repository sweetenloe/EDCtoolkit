#Requires -Version 5.1
<#
EDC Toolkit (Beta GUI)

Run:
  powershell.exe -ExecutionPolicy Bypass -File .\EDCtoolkit.GUI.ps1

Notes:
  - Run as Administrator for full visibility and to apply some fixes.
  - This GUI wraps and reuses reporting/helpers from `edctoolkit.ps1`.
#>

[CmdletBinding()]
param(
    [switch]$SelfTest
    ,
    [switch]$SelfTestFast
    ,
    [ValidateSet('System','Dark','Light')]
    [string]$Theme = 'System'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Warning 'WinForms requires STA. Re-run using Windows PowerShell (powershell.exe), not pwsh.exe.'
}

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# -----------------------------
# Load existing toolkit backend
# -----------------------------

$script:ToolkitPath = Join-Path -Path $PSScriptRoot -ChildPath 'edctoolkit.ps1'
if (-not (Test-Path -Path $script:ToolkitPath)) {
    [System.Windows.Forms.MessageBox]::Show("Missing backend script: $script:ToolkitPath", 'EDC Toolkit Beta', 'OK', 'Error') | Out-Null
    return
}

. $script:ToolkitPath

try {
    Initialize-ReportSession -Name ('GUI_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss')) -NonInteractive
}
catch {
    # If report session creation fails, the GUI still runs; exporting will fall back to direct file writes.
}

# -----------------------------
# Normalized result model
# -----------------------------

function New-EdcResult {
    param(
        [Parameter(Mandatory)][ValidateSet('System','Network','File System','Users','Security','Services','Tools / Utilities','Reports / Logs')]
        [string]$Category,
        [Parameter(Mandatory)][string]$CheckName,
        [Parameter(Mandatory)][ValidateSet('OK','Warning','Failed','Not available')]
        [string]$Status,
        [Parameter(Mandatory)][string]$Summary,
        [Parameter(Mandatory)][string]$Details,
        [string]$ResultGroup = 'Health Checks',
        [string]$RecommendedFix = '',
        [bool]$CanAutoFix = $false,
        [AllowNull()][scriptblock]$FixAction = $null,
        [bool]$FixRequiresAdmin = $false,
        [bool]$FixDisruptive = $false,
        [string]$ReportPath = ''
    )

    [pscustomobject]@{
        Category         = $Category
        CheckName        = $CheckName
        Status           = $Status
        Summary          = $Summary
        Details          = $Details
        ResultGroup      = $ResultGroup
        RecommendedFix   = $RecommendedFix
        CanAutoFix       = $CanAutoFix
        FixAction        = $FixAction
        FixRequiresAdmin = $FixRequiresAdmin
        FixDisruptive    = $FixDisruptive
        ReportPath       = $ReportPath
        CollectedAt      = Get-Date
    }
}

function Get-ResultKey {
    param([Parameter(Mandatory)][object]$Result)
    '{0}|{1}' -f $Result.Category, $Result.CheckName
}

$script:ResultsByKey = @{}
$script:TopCategoryButtons = @()

function Merge-Results {
    param([Parameter(Mandatory)][object[]]$Results)
    foreach ($r in $Results) {
        $script:ResultsByKey[(Get-ResultKey -Result $r)] = $r
    }
}

function Get-AllResults {
    $script:ResultsByKey.Values | Sort-Object Category, CheckName
}

function Get-CategoryResults {
    param([Parameter(Mandatory)][string]$Category)
    $script:ResultsByKey.Values | Where-Object { $_.Category -eq $Category } | Sort-Object CheckName
}

# -----------------------------
# Status visuals + safe helpers
# -----------------------------

function New-StatusImageList {
    $imgList = New-Object System.Windows.Forms.ImageList
    $imgList.ImageSize = New-Object System.Drawing.Size(16,16)
    $imgList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit

    function New-StatusBitmap {
        param(
            [Parameter(Mandatory)][System.Drawing.Color]$BackColor,
            [Parameter(Mandatory)][string]$Glyph
        )
        $bmp = New-Object System.Drawing.Bitmap 16,16
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $g.Clear([System.Drawing.Color]::Transparent)
            $brush = New-Object System.Drawing.SolidBrush($BackColor)
            $g.FillEllipse($brush, 1, 1, 14, 14)
            $brush.Dispose()

            $font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            $sf = New-Object System.Drawing.StringFormat
            $sf.Alignment = [System.Drawing.StringAlignment]::Center
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
            $g.DrawString($Glyph, $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0,0,16,16)), $sf)
            $font.Dispose()
            $sf.Dispose()
            return $bmp
        }
        finally {
            $g.Dispose()
        }
    }

    $imgList.Images.Add('OK', (New-StatusBitmap -BackColor ([System.Drawing.Color]::FromArgb(46, 160, 67)) -Glyph '✓')) | Out-Null
    $imgList.Images.Add('Warning', (New-StatusBitmap -BackColor ([System.Drawing.Color]::FromArgb(245, 158, 11)) -Glyph '!')) | Out-Null
    $imgList.Images.Add('Failed', (New-StatusBitmap -BackColor ([System.Drawing.Color]::FromArgb(220, 38, 38)) -Glyph '×')) | Out-Null
    $imgList.Images.Add('Not available', (New-StatusBitmap -BackColor ([System.Drawing.Color]::FromArgb(107, 114, 128)) -Glyph '?')) | Out-Null
    return $imgList
}

function Get-StatusImageKey {
    param([Parameter(Mandatory)][string]$Status)
    switch ($Status) {
        'OK' { 'OK' }
        'Warning' { 'Warning' }
        'Failed' { 'Failed' }
        default { 'Not available' }
    }
}

function Invoke-EdcSafe {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )
    try { & $ScriptBlock } catch { return $null }
}

function Get-IsAdmin {
    try { return (Test-IsAdmin) } catch { return $false }
}

function Confirm-GuiAction {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Title = 'EDC Toolkit Beta'
    )
    $result = [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Get-ThemePreference {
    param(
        [ValidateSet('System','Dark','Light')]
        [string]$Default = 'System'
    )

    if ($Default -eq 'Dark' -or $Default -eq 'Light') { return $Default }

    try {
        $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        $value = Get-ItemPropertyValue -Path $key -Name 'AppsUseLightTheme' -ErrorAction Stop
        if ([int]$value -eq 0) { return 'Dark' }
        return 'Light'
    }
    catch {
        return 'Dark'
    }
}

function New-ThemePalette {
    param(
        [Parameter(Mandatory)][ValidateSet('Dark','Light')]
        [string]$Theme
    )

    if ($Theme -eq 'Light') {
        return @{
            ThemeName      = 'Light'
            FormBack       = [System.Drawing.Color]::FromArgb(245, 247, 250)
            Surface        = [System.Drawing.Color]::White
            SurfaceAlt     = [System.Drawing.Color]::FromArgb(237, 242, 247)
            Border         = [System.Drawing.Color]::FromArgb(210, 214, 220)
            Text           = [System.Drawing.Color]::FromArgb(17, 24, 39)
            MutedText      = [System.Drawing.Color]::FromArgb(75, 85, 99)
            InputBack      = [System.Drawing.Color]::White
            InputText      = [System.Drawing.Color]::FromArgb(17, 24, 39)
            ButtonBack     = [System.Drawing.Color]::FromArgb(232, 236, 241)
            ButtonText     = [System.Drawing.Color]::FromArgb(17, 24, 39)
            Accent         = [System.Drawing.Color]::FromArgb(37, 99, 235)
            StatusBack     = [System.Drawing.Color]::FromArgb(248, 250, 252)
            ListBack       = [System.Drawing.Color]::White
            ListFore       = [System.Drawing.Color]::FromArgb(17, 24, 39)
        }
    }

    return @{
        ThemeName      = 'Dark'
        FormBack       = [System.Drawing.Color]::FromArgb(18, 23, 31)
        Surface        = [System.Drawing.Color]::FromArgb(28, 35, 46)
        SurfaceAlt     = [System.Drawing.Color]::FromArgb(35, 43, 56)
        Border         = [System.Drawing.Color]::FromArgb(56, 66, 82)
        Text           = [System.Drawing.Color]::FromArgb(232, 237, 243)
        MutedText      = [System.Drawing.Color]::FromArgb(161, 173, 189)
        InputBack      = [System.Drawing.Color]::FromArgb(23, 29, 39)
        InputText      = [System.Drawing.Color]::FromArgb(232, 237, 243)
        ButtonBack     = [System.Drawing.Color]::FromArgb(42, 50, 64)
        ButtonText     = [System.Drawing.Color]::FromArgb(232, 237, 243)
        Accent         = [System.Drawing.Color]::FromArgb(89, 168, 255)
        StatusBack     = [System.Drawing.Color]::FromArgb(24, 31, 41)
        ListBack       = [System.Drawing.Color]::FromArgb(23, 29, 39)
        ListFore       = [System.Drawing.Color]::FromArgb(232, 237, 243)
    }
}

function Set-ListViewTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.ListView]$ListView,
        [Parameter(Mandatory)][hashtable]$Palette
    )

    $ListView.BackColor = $Palette.ListBack
    $ListView.ForeColor = $Palette.ListFore
}

function Set-RoundedButtonRegion {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [int]$Radius = 14
    )

    if ($Button.Width -lt 8 -or $Button.Height -lt 8) { return }

    $diameter = [math]::Max(4, ($Radius * 2))
    $rect = New-Object System.Drawing.Rectangle(0, 0, $Button.Width, $Button.Height)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    try {
        $path.StartFigure()
        $path.AddArc($rect.X, $rect.Y, $diameter, $diameter, 180, 90)
        $path.AddArc(($rect.Right - $diameter), $rect.Y, $diameter, $diameter, 270, 90)
        $path.AddArc(($rect.Right - $diameter), ($rect.Bottom - $diameter), $diameter, $diameter, 0, 90)
        $path.AddArc($rect.X, ($rect.Bottom - $diameter), $diameter, $diameter, 90, 90)
        $path.CloseFigure()

        if ($Button.Region) {
            try { $Button.Region.Dispose() } catch { }
        }
        $Button.Region = New-Object System.Drawing.Region($path)
    }
    finally {
        $path.Dispose()
    }
}

function Set-RichTextTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.RichTextBox]$RichTextBox,
        [Parameter(Mandatory)][hashtable]$Palette
    )

    $RichTextBox.BackColor = $Palette.InputBack
    $RichTextBox.ForeColor = $Palette.InputText
    $RichTextBox.BorderStyle = 'FixedSingle'
}

function Set-ComboTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.ComboBox]$ComboBox,
        [Parameter(Mandatory)][hashtable]$Palette
    )

    $ComboBox.BackColor = $Palette.InputBack
    $ComboBox.ForeColor = $Palette.InputText
    $ComboBox.FlatStyle = 'Flat'
    $ComboBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $ComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    $ComboBox.add_DrawItem({
        param($sender, $e)
        if ($e.Index -lt 0) { return }

        $item = $sender.Items[$e.Index]
        $text = if ($item -and $item.PSObject.Properties['Text']) { [string]$item.Text } else { [string]$item }
        $isSelected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected
        $back = if ($isSelected) { $script:Palette.Accent } else { $script:Palette.InputBack }
        $fore = if ($isSelected) { [System.Drawing.Color]::White } else { $script:Palette.InputText }

        $bg = New-Object System.Drawing.SolidBrush($back)
        $fg = New-Object System.Drawing.SolidBrush($fore)
        try {
            $e.Graphics.FillRectangle($bg, $e.Bounds)
            $textRect = New-Object System.Drawing.RectangleF(($e.Bounds.X + 6), ($e.Bounds.Y + 3), ($e.Bounds.Width - 8), ($e.Bounds.Height - 4))
            $e.Graphics.DrawString($text, $sender.Font, $fg, $textRect)
            $e.DrawFocusRectangle()
        }
        finally {
            $bg.Dispose()
            $fg.Dispose()
        }
    }.GetNewClosure())
}

function Set-TabControlTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TabControl]$TabControl,
        [Parameter(Mandatory)][hashtable]$Palette
    )

    $TabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $TabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Normal
    $TabControl.Appearance = [System.Windows.Forms.TabAppearance]::Normal
    $TabControl.Padding = New-Object System.Drawing.Point(14, 6)

    $TabControl.add_DrawItem({
        param($sender, $e)
        if ($e.Index -lt 0) { return }

        $tabPage = $sender.TabPages[$e.Index]
        $rect = $sender.GetTabRect($e.Index)
        $selected = ($sender.SelectedIndex -eq $e.Index)
        $back = if ($selected) { $script:Palette.Surface } else { $script:Palette.SurfaceAlt }
        $fore = if ($selected) { $script:Palette.Text } else { $script:Palette.MutedText }

        $bg = New-Object System.Drawing.SolidBrush($back)
        $fg = New-Object System.Drawing.SolidBrush($fore)
        $borderPen = New-Object System.Drawing.Pen($(if ($selected) { $script:Palette.Accent } else { $script:Palette.Border }))
        $sf = New-Object System.Drawing.StringFormat
        try {
            $sf.Alignment = [System.Drawing.StringAlignment]::Center
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
            $e.Graphics.FillRectangle($bg, $rect)
            $e.Graphics.DrawRectangle($borderPen, $rect)
            $textRect = New-Object System.Drawing.RectangleF($rect.X, $rect.Y, $rect.Width, $rect.Height)
            $e.Graphics.DrawString($tabPage.Text, $sender.Font, $fg, $textRect, $sf)
        }
        finally {
            $bg.Dispose()
            $fg.Dispose()
            $borderPen.Dispose()
            $sf.Dispose()
        }
    }.GetNewClosure())
}

function Set-ListViewOwnerTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.ListView]$ListView,
        [Parameter(Mandatory)][hashtable]$Palette
    )

    $ListView.OwnerDraw = $true
    $ListView.BackColor = $Palette.ListBack
    $ListView.ForeColor = $Palette.ListFore

    $ListView.add_DrawColumnHeader({
        param($sender, $e)
        $bg = New-Object System.Drawing.SolidBrush($script:Palette.SurfaceAlt)
        $fg = New-Object System.Drawing.SolidBrush($script:Palette.Text)
        $pen = New-Object System.Drawing.Pen($script:Palette.Border)
        $sf = New-Object System.Drawing.StringFormat
        try {
            $sf.Alignment = [System.Drawing.StringAlignment]::Near
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
            $e.Graphics.FillRectangle($bg, $e.Bounds)
            $e.Graphics.DrawRectangle($pen, $e.Bounds)
            $textRect = New-Object System.Drawing.RectangleF(($e.Bounds.X + 6), $e.Bounds.Y, ($e.Bounds.Width - 8), $e.Bounds.Height)
            $e.Graphics.DrawString($e.Header.Text, $sender.Font, $fg, $textRect, $sf)
        }
        finally {
            $bg.Dispose()
            $fg.Dispose()
            $pen.Dispose()
            $sf.Dispose()
        }
    }.GetNewClosure())

    $ListView.add_DrawItem({
        param($sender, $e)
        if ($sender.View -ne [System.Windows.Forms.View]::Details) {
            $e.DrawDefault = $true
        }
    }.GetNewClosure())

    $ListView.add_DrawSubItem({
        param($sender, $e)
        $isSelected = ($e.ItemState -band [System.Windows.Forms.ListViewItemStates]::Selected) -eq [System.Windows.Forms.ListViewItemStates]::Selected
        $rowBack = if ($isSelected) { $script:Palette.Accent } elseif (($e.ItemIndex % 2) -eq 0) { $script:Palette.ListBack } else { $script:Palette.Surface }
        $rowFore = if ($isSelected) { [System.Drawing.Color]::White } else { $script:Palette.ListFore }

        $bg = New-Object System.Drawing.SolidBrush($rowBack)
        $fg = New-Object System.Drawing.SolidBrush($rowFore)
        $gridPen = New-Object System.Drawing.Pen($script:Palette.Border)
        $sf = New-Object System.Drawing.StringFormat
        try {
            $sf.Alignment = [System.Drawing.StringAlignment]::Near
            $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
            $e.Graphics.FillRectangle($bg, $e.Bounds)
            $e.Graphics.DrawRectangle($gridPen, $e.Bounds)
            $textRect = New-Object System.Drawing.RectangleF(($e.Bounds.X + 6), $e.Bounds.Y, ($e.Bounds.Width - 8), $e.Bounds.Height)
            $e.Graphics.DrawString($e.SubItem.Text, $sender.Font, $fg, $textRect, $sf)

            if ($e.ColumnIndex -eq 0 -and $e.Item.ImageList -and $e.Item.ImageIndex -ge 0) {
                $e.Item.ImageList.Draw($e.Graphics, ($e.Bounds.X + 4), ($e.Bounds.Y + 2), $e.Item.ImageIndex)
            }
        }
        finally {
            $bg.Dispose()
            $fg.Dispose()
            $gridPen.Dispose()
            $sf.Dispose()
        }
    }.GetNewClosure())
}

function Apply-ControlTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Control,
        [Parameter(Mandatory)][hashtable]$Palette
    )

    if ($Control -is [System.Windows.Forms.Form]) {
        $Control.BackColor = $Palette.FormBack
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.StatusStrip]) {
        $Control.BackColor = $Palette.StatusBack
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.TabPage]) {
        $Control.BackColor = $Palette.Surface
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.GroupBox]) {
        $Control.BackColor = $Palette.Surface
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.Button]) {
        $Control.BackColor = $Palette.ButtonBack
        $Control.ForeColor = $Palette.ButtonText
        $Control.FlatStyle = 'Flat'
        $Control.FlatAppearance.BorderSize = 1
        $Control.FlatAppearance.BorderColor = $Palette.Accent
        $Control.FlatAppearance.MouseDownBackColor = $Palette.Accent
        $Control.FlatAppearance.MouseOverBackColor = $Palette.SurfaceAlt
        $Control.UseVisualStyleBackColor = $false
        $Control.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
        Set-RoundedButtonRegion -Button $Control -Radius 12
        $Control.add_SizeChanged({
            Set-RoundedButtonRegion -Button $this -Radius 12
        }.GetNewClosure())
    }
    elseif ($Control -is [System.Windows.Forms.TextBox] -or $Control -is [System.Windows.Forms.ComboBox]) {
        if ($Control -is [System.Windows.Forms.ComboBox]) {
            Set-ComboTheme -ComboBox $Control -Palette $Palette
        }
        else {
            $Control.BackColor = $Palette.InputBack
            $Control.ForeColor = $Palette.InputText
            $Control.BorderStyle = 'FixedSingle'
        }
    }
    elseif ($Control -is [System.Windows.Forms.Label]) {
        if ($Control.Name -eq 'SubtitleLabel') {
            $Control.ForeColor = $Palette.MutedText
        }
        else {
            $Control.ForeColor = $Palette.Text
        }
        if ($Control.BackColor -eq [System.Drawing.Color]::Empty) {
            $Control.BackColor = [System.Drawing.Color]::Transparent
        }
    }
    elseif ($Control -is [System.Windows.Forms.Panel] -or $Control -is [System.Windows.Forms.SplitContainer]) {
        if ($Control.Name -eq 'SeparatorPanel') {
            $Control.BackColor = $Palette.Border
        }
        elseif ($Control.Name -eq 'HeaderPanel') {
            $Control.BackColor = $Palette.FormBack
        }
        else {
            $Control.BackColor = $Palette.Surface
            $Control.ForeColor = $Palette.Text
        }
    }
    elseif ($Control -is [System.Windows.Forms.ListView]) {
        Set-ListViewTheme -ListView $Control -Palette $Palette
        Set-ListViewOwnerTheme -ListView $Control -Palette $Palette
    }
    elseif ($Control -is [System.Windows.Forms.RichTextBox]) {
        Set-RichTextTheme -RichTextBox $Control -Palette $Palette
    }
    elseif ($Control -is [System.Windows.Forms.TabControl]) {
        $Control.BackColor = $Palette.FormBack
        $Control.ForeColor = $Palette.Text
        Set-TabControlTheme -TabControl $Control -Palette $Palette
    }
    else {
        $Control.ForeColor = $Palette.Text
    }

    foreach ($child in $Control.Controls) {
        if ($child -is [System.Windows.Forms.Control]) {
            Apply-ControlTheme -Control $child -Palette $Palette
        }
    }
}

# -----------------------------
# Toolkit report capture helpers
# -----------------------------

function Try-GetLatestNewReportText {
    param(
        [Parameter(Mandatory)][datetime]$Baseline,
        [int]$MaxChars = 200000
    )
    try {
        if (-not (Get-Variable -Name ReportRoot -Scope Script -ErrorAction SilentlyContinue)) { return $null }
        $root = $Script:ReportRoot
        if (-not $root -or -not (Test-Path -Path $root)) { return $null }

        $file = Get-ChildItem -Path $root -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt $Baseline } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if (-not $file) { return $null }

        $text = Get-Content -Path $file.FullName -Raw -ErrorAction SilentlyContinue
        if ($null -eq $text) { return $null }
        if ($text.Length -gt $MaxChars) { return ($text.Substring(0, $MaxChars) + "`r`n...[truncated]") }
        return $text
    }
    catch {
        return $null
    }
}

function Invoke-ToolkitHandlersAsInventory {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string[]]$Handlers
    )

    $baseline = Get-Date
    $chunks = New-Object System.Collections.Generic.List[string]
    foreach ($h in $Handlers) {
        if (-not (Get-Command -Name $h -CommandType Function -ErrorAction SilentlyContinue)) {
            $chunks.Add(("=== {0} ===`r`nNot available (missing function).`r`n" -f $h)) | Out-Null
            continue
        }

        $null = Invoke-EdcSafe -Name $h -ScriptBlock { & $h }
        $report = Try-GetLatestNewReportText -Baseline $baseline
        $baseline = Get-Date

        if ([string]::IsNullOrWhiteSpace($report)) {
            $chunks.Add(("=== {0} ===`r`nNo report output captured.`r`n" -f $h)) | Out-Null
        }
        else {
            $chunks.Add(("=== {0} ===`r`n{1}`r`n" -f $h, $report.Trim())) | Out-Null
        }
    }

    $text = ($chunks.ToArray() -join "`r`n")
    $hasReal = $text -match '[^\s]'
    New-EdcResult -Category $Category -CheckName 'Toolkit Inventory' -Status $(if ($hasReal) { 'OK' } else { 'Not available' }) `
        -ResultGroup 'Toolkit Output' `
        -Summary 'Collected toolkit outputs for this category.' `
        -Details $text `
        -RecommendedFix '' -CanAutoFix:$false -FixAction $null
}

# -----------------------------
# Report browsing (Reports/Logs tab)
# -----------------------------

function Get-CurrentReportRoot {
    try {
        if (Get-Variable -Name ReportRoot -Scope Script -ErrorAction SilentlyContinue) {
            if ($Script:ReportRoot -and (Test-Path -Path $Script:ReportRoot)) { return $Script:ReportRoot }
        }
    }
    catch { }

    try {
        if (Get-Command -Name Get-DefaultReportBaseRoot -CommandType Function -ErrorAction SilentlyContinue) {
            return (Get-DefaultReportBaseRoot)
        }
    }
    catch { }

    return (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'EDCtoolkit\EDC_Reports')
}

function Get-TextFilePreview {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxChars = 200000
    )
    try {
        $text = Get-Content -Path $Path -Raw -ErrorAction SilentlyContinue
        if ($null -eq $text) { return '' }
        if ($text.Length -le $MaxChars) { return $text }
        return ($text.Substring(0, $MaxChars) + "`r`n...[truncated]")
    }
    catch {
        return ("Unable to read file: {0}" -f $_.Exception.Message)
    }
}

function Sync-ReportFilesToResults {
    param(
        [int]$MaxFiles = 250,
        [int]$MaxChars = 200000
    )

    $root = Get-CurrentReportRoot
    if (-not (Test-Path -Path $root)) {
        $na = New-EdcResult -Category 'Reports / Logs' -CheckName 'Reports folder' -Status 'Not available' -Summary 'Reports folder not found.' -Details $root -ResultGroup 'Report Files'
        Merge-Results -Results @($na)
        return
    }

    $files = @(Get-ChildItem -Path $root -File -Filter '*.txt' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First $MaxFiles)
    $folderSummary = "Root: $root`r`nFiles: $($files.Count)"
    $folderResult = New-EdcResult -Category 'Reports / Logs' -CheckName 'Reports folder' -Status 'OK' -Summary ("{0} text reports" -f $files.Count) -Details $folderSummary -ResultGroup 'Report Files' -ReportPath $root
    Merge-Results -Results @($folderResult)

    foreach ($f in $files) {
        $kb = [math]::Round(($f.Length / 1KB), 1)
        $summary = "{0}  ({1} KB)" -f $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'), $kb
        $details = Get-TextFilePreview -Path $f.FullName -MaxChars $MaxChars
        $r = New-EdcResult -Category 'Reports / Logs' -CheckName ("Report: {0}" -f $f.Name) -Status 'OK' -Summary $summary -Details $details -ResultGroup 'Report Files' -ReportPath $f.FullName
        Merge-Results -Results @($r)
    }
}

# -----------------------------
# Scan + troubleshoot checks
# -----------------------------

$script:PingTarget = '8.8.8.8'
$script:DnsTarget = 'www.microsoft.com'
$script:ResolvedTheme = Get-ThemePreference -Default $Theme
$script:Palette = New-ThemePalette -Theme $script:ResolvedTheme

function Get-SystemChecks {
    $results = New-Object System.Collections.Generic.List[object]

    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($null -eq $os) {
        $results.Add((New-EdcResult -Category 'System' -CheckName 'OS info readable' -Status 'Not available' -Summary 'Unable to read OS info.' -Details 'Get-CimInstance Win32_OperatingSystem returned no data.' -RecommendedFix 'Run as Administrator and ensure WMI/CIM is functional.' )) | Out-Null
    }
    else {
        $caption = [string]$os.Caption
        $ver = [string]$os.Version
        $build = [string]$os.BuildNumber
        $arch = [string]$os.OSArchitecture
        $results.Add((New-EdcResult -Category 'System' -CheckName 'OS info readable' -Status 'OK' -Summary ("{0} {1} (Build {2}, {3})" -f $caption, $ver, $build, $arch) -Details ($os | Format-List * | Out-String -Width 4096).Trim())) | Out-Null
    }

    $lastBoot = $null
    if ($os) { $lastBoot = Convert-FromDmtfDateTime -Value $os.LastBootUpTime -Quiet }
    if ($lastBoot) {
        $uptime = (Get-Date) - $lastBoot
        $days = [math]::Round($uptime.TotalDays, 2)
        $status = if ($days -ge 30) { 'Warning' } else { 'OK' }
        $why = if ($status -eq 'Warning') { 'Unusually long uptime can correlate with pending updates, memory pressure, or stale network state.' } else { 'Uptime looks normal.' }
        $results.Add((New-EdcResult -Category 'System' -CheckName 'Uptime' -Status $status -Summary ("{0} days (Last boot: {1})" -f $days, $lastBoot.ToString('yyyy-MM-dd HH:mm:ss')) -Details $why -RecommendedFix 'Consider scheduling a reboot during a maintenance window if issues are observed.' )) | Out-Null
    }
    else {
        $results.Add((New-EdcResult -Category 'System' -CheckName 'Uptime' -Status 'Not available' -Summary 'Unable to determine uptime.' -Details 'LastBootUpTime could not be parsed.' -RecommendedFix 'Ensure WMI/CIM is accessible.' )) | Out-Null
    }

    $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
    if (-not $disks) {
        $results.Add((New-EdcResult -Category 'System' -CheckName 'Disk free space' -Status 'Not available' -Summary 'No fixed disks returned.' -Details 'Win32_LogicalDisk query returned no data.' -RecommendedFix 'Ensure WMI/CIM is accessible.' )) | Out-Null
    }
    else {
        $lines = New-Object System.Collections.Generic.List[string]
        $worst = 'OK'
        foreach ($d in $disks) {
            if (-not $d.Size -or -not $d.FreeSpace) {
                $lines.Add(("{0}: size/free unavailable" -f $d.DeviceID)) | Out-Null
                $worst = if ($worst -eq 'OK') { 'Warning' } else { $worst }
                continue
            }
            $pctFree = [math]::Round(($d.FreeSpace / $d.Size) * 100, 2)
            $freeGb = [math]::Round(($d.FreeSpace / 1GB), 2)
            $sizeGb = [math]::Round(($d.Size / 1GB), 2)
            $lines.Add(("{0} ({1}): {2}% free ({3} GB / {4} GB)" -f $d.DeviceID, $d.VolumeName, $pctFree, $freeGb, $sizeGb)) | Out-Null
            if ($pctFree -lt 10) { $worst = 'Failed' }
            elseif ($pctFree -lt 15 -and $worst -ne 'Failed') { $worst = 'Warning' }
        }

        $fix = ''
        $canFix = $false
        $fixAction = $null
        if ($worst -ne 'OK') {
            $fix = 'Free space by removing temporary files and large unused data; consider moving data off the drive.'
            $canFix = $true
            $fixAction = {
                $temp = $env:TEMP
                if ([string]::IsNullOrWhiteSpace($temp) -or -not (Test-Path -Path $temp)) { return }
                Get-ChildItem -Path $temp -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        $summary = if ($lines.Count -gt 0) { $lines[0] } else { 'Disk space information available.' }
        $results.Add((New-EdcResult -Category 'System' -CheckName 'Disk free space' -Status $worst -Summary $summary -Details (($lines.ToArray() -join "`r`n")) -RecommendedFix $fix -CanAutoFix:$canFix -FixAction $fixAction -FixRequiresAdmin:$false -FixDisruptive:$true)) | Out-Null
    }

    if ($os) {
        $totalKb = [double]$os.TotalVisibleMemorySize
        $freeKb = [double]$os.FreePhysicalMemory
        if ($totalKb -gt 0) {
            $pctFree = [math]::Round(($freeKb / $totalKb) * 100, 2)
            $freeGb = [math]::Round(($freeKb * 1KB) / 1GB, 2)
            $totalGb = [math]::Round(($totalKb * 1KB) / 1GB, 2)
            $status = if ($pctFree -lt 5) { 'Failed' } elseif ($pctFree -lt 10) { 'Warning' } else { 'OK' }
            $results.Add((New-EdcResult -Category 'System' -CheckName 'RAM availability' -Status $status -Summary ("{0}% free ({1} GB / {2} GB)" -f $pctFree, $freeGb, $totalGb) -Details 'Low free RAM can indicate memory pressure or heavy background load.' -RecommendedFix 'Close heavy applications; consider rebooting if memory usage remains high.' )) | Out-Null
        }
        else {
            $results.Add((New-EdcResult -Category 'System' -CheckName 'RAM availability' -Status 'Not available' -Summary 'Memory counters unavailable.' -Details 'TotalVisibleMemorySize returned 0.' -RecommendedFix 'Ensure WMI/CIM is accessible.' )) | Out-Null
        }
    }

    $start = (Get-Date).AddHours(-24)
    $evtCount = $null
    try {
        $events = Get-WinEvent -FilterHashtable @{ LogName='System'; Level=2; StartTime=$start } -ErrorAction SilentlyContinue
        if ($events) { $evtCount = @($events).Count } else { $evtCount = 0 }
    }
    catch { $evtCount = $null }

    if ($null -eq $evtCount) {
        $results.Add((New-EdcResult -Category 'System' -CheckName 'Event log errors (24h)' -Status 'Not available' -Summary 'Unable to query System event log.' -Details 'Get-WinEvent failed or is blocked.' -RecommendedFix 'Run as Administrator or verify Event Log service.' )) | Out-Null
    }
    else {
        $status = if ($evtCount -ge 500) { 'Failed' } elseif ($evtCount -ge 100) { 'Warning' } else { 'OK' }
        $summary = "{0} System error events (Level=2) in last 24h" -f $evtCount
        $details = 'Event log error volume is noisy by nature; this is a conservative warning signal.'
        $results.Add((New-EdcResult -Category 'System' -CheckName 'Event log errors (24h)' -Status $status -Summary $summary -Details $details -RecommendedFix 'Review recent errors for repeating providers/IDs; fix underlying driver or service issues if patterns emerge.' )) | Out-Null
    }

    return $results.ToArray()
}

function Get-NetworkChecks {
    param(
        [string]$PingTarget,
        [string]$DnsTarget = 'www.microsoft.com'
    )
    $results = New-Object System.Collections.Generic.List[object]

    $hasIp = $false
    $hasGw = $false
    $adapterSummary = New-Object System.Collections.Generic.List[string]

    try {
        if (Get-Command -Name Get-NetIPConfiguration -ErrorAction SilentlyContinue) {
            $cfg = Get-NetIPConfiguration -ErrorAction SilentlyContinue
            foreach ($c in $cfg) {
                $ip4 = $c.IPv4Address | Select-Object -First 1
                $gw4 = $c.IPv4DefaultGateway | Select-Object -First 1
                $line = "{0}: {1} / GW {2}" -f $c.InterfaceAlias, $(if ($ip4) { $ip4.IPAddress } else { 'no IPv4' }), $(if ($gw4) { $gw4.NextHop } else { 'none' })
                $adapterSummary.Add($line) | Out-Null
                if ($ip4) { $hasIp = $true }
                if ($gw4) { $hasGw = $true }
            }
        }
        else {
            $cfg = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction SilentlyContinue
            foreach ($c in $cfg) {
                $ip = ($c.IPAddress | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
                $gw = ($c.DefaultIPGateway | Where-Object { $_ -match '^\d+\.' } | Select-Object -First 1)
                $adapterSummary.Add(("{0}: {1} / GW {2}" -f $c.Description, $(if ($ip) { $ip } else { 'no IPv4' }), $(if ($gw) { $gw } else { 'none' }))) | Out-Null
                if ($ip) { $hasIp = $true }
                if ($gw) { $hasGw = $true }
            }
        }
    }
    catch { }

    if ($adapterSummary.Count -eq 0) {
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'Active adapter present' -Status 'Not available' -Summary 'No adapter configuration data returned.' -Details 'Unable to read adapter/IP configuration via Get-NetIPConfiguration or WMI.' -RecommendedFix 'Ensure network stack and WMI are functional; run as Administrator for best results.' )) | Out-Null
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'IP assigned' -Status 'Not available' -Summary 'Unable to determine if an IP is assigned.' -Details 'Adapter/IP configuration was not readable.' -RecommendedFix '' )) | Out-Null
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'Default gateway present' -Status 'Not available' -Summary 'Unable to determine default gateway.' -Details 'Adapter/IP configuration was not readable.' -RecommendedFix '' )) | Out-Null
    }
    else {
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'Active adapter present' -Status 'OK' -Summary 'Adapter configuration readable.' -Details (($adapterSummary.ToArray() -join "`r`n")) -RecommendedFix '' )) | Out-Null
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'IP assigned' -Status $(if ($hasIp) { 'OK' } else { 'Failed' }) -Summary $(if ($hasIp) { 'At least one interface has an IPv4 address.' } else { 'No IPv4 address detected on IP-enabled interfaces.' }) -Details (($adapterSummary.ToArray() -join "`r`n")) -RecommendedFix $(if ($hasIp) { '' } else { 'Verify cable/Wi-Fi association, DHCP, and adapter state.' }) )) | Out-Null
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'Default gateway present' -Status $(if ($hasGw) { 'OK' } else { 'Warning' }) -Summary $(if ($hasGw) { 'Default gateway detected.' } else { 'No default gateway detected (may be normal on isolated networks).' }) -Details (($adapterSummary.ToArray() -join "`r`n")) -RecommendedFix $(if ($hasGw) { '' } else { 'If internet access is expected, verify DHCP and routing configuration.' }) )) | Out-Null
    }

    $dnsOk = $null
    $dnsDetails = ''
    $dnsName = if ([string]::IsNullOrWhiteSpace($DnsTarget)) { 'www.microsoft.com' } else { $DnsTarget.Trim() }
    try {
        if (Get-Command -Name Resolve-DnsName -ErrorAction SilentlyContinue) {
            $r = Resolve-DnsName -Name $dnsName -ErrorAction SilentlyContinue | Select-Object -First 1
            $dnsOk = [bool]$r
        }
        else {
            $entry = [System.Net.Dns]::GetHostEntry($dnsName)
            $dnsOk = ($entry -and $entry.AddressList -and $entry.AddressList.Count -gt 0)
        }
        $dnsDetails = "DNS resolution for $dnsName succeeded."
    }
    catch {
        $dnsOk = $false
        $dnsDetails = $_.Exception.Message
    }

    if ($null -eq $dnsOk) {
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'DNS resolution test' -Status 'Not available' -Summary 'DNS test not available.' -Details 'Resolve-DnsName and fallback resolution were unavailable.' -RecommendedFix 'Verify DNS client service and name resolution settings.' )) | Out-Null
    }
    elseif ($dnsOk) {
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'DNS resolution test' -Status 'OK' -Summary "DNS resolution succeeded for $dnsName." -Details $dnsDetails -RecommendedFix '' )) | Out-Null
    }
    else {
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'DNS resolution test' -Status 'Failed' -Summary "DNS resolution failed for $dnsName." -Details $dnsDetails -RecommendedFix 'Check DNS server settings and network connectivity.' -CanAutoFix:$true -FixAction { ipconfig /flushdns | Out-Null } -FixRequiresAdmin:$false -FixDisruptive:$false)) | Out-Null
    }

    $target = $PingTarget
    if (-not [string]::IsNullOrWhiteSpace($target)) {
        $pingOk = $null
        $msg = ''
        try {
            $pingOk = Test-Connection -ComputerName $target -Count 1 -Quiet -ErrorAction SilentlyContinue
            $msg = "Ping to $target returned: $pingOk"
        }
        catch {
            $pingOk = $false
            $msg = $_.Exception.Message
        }

        if ($null -eq $pingOk) {
            $results.Add((New-EdcResult -Category 'Network' -CheckName 'Ping test' -Status 'Not available' -Summary 'Ping test not available.' -Details 'Test-Connection returned no result.' -RecommendedFix '' )) | Out-Null
        }
        elseif ($pingOk) {
            $results.Add((New-EdcResult -Category 'Network' -CheckName 'Ping test' -Status 'OK' -Summary "Ping to $target succeeded." -Details $msg -RecommendedFix '' )) | Out-Null
        }
        else {
            $results.Add((New-EdcResult -Category 'Network' -CheckName 'Ping test' -Status 'Warning' -Summary "Ping to $target failed (may be blocked by policy)." -Details $msg -RecommendedFix 'If connectivity issues are suspected, try a different target or verify ICMP is allowed by policy/firewall.' )) | Out-Null
        }
    }

    $ports = $null
    try {
        if (Get-Command -Name Get-NetTCPConnection -ErrorAction SilentlyContinue) {
            $ports = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Select-Object -First 200 LocalAddress,LocalPort,OwningProcess
        }
    }
    catch { $ports = $null }

    if ($ports) {
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'Listening ports summary' -Status 'OK' -Summary ("{0} listening sockets (sampled)" -f @($ports).Count) -Details (($ports | Format-Table -AutoSize | Out-String -Width 4096).Trim()) -RecommendedFix '' )) | Out-Null
    }
    else {
        $results.Add((New-EdcResult -Category 'Network' -CheckName 'Listening ports summary' -Status 'Not available' -Summary 'Listening port query not available.' -Details 'Get-NetTCPConnection is unavailable or returned no data.' -RecommendedFix '' )) | Out-Null
    }

    return $results.ToArray()
}

function Get-UsersChecks {
    $results = New-Object System.Collections.Generic.List[object]

    $adminsText = $null
    try {
        if (Get-Command -Name Get-LocalGroupMember -ErrorAction SilentlyContinue) {
            $admins = Get-LocalGroupMember -Group 'Administrators' -ErrorAction SilentlyContinue | Select-Object Name, ObjectClass, PrincipalSource
            $adminsText = ($admins | Format-Table -AutoSize | Out-String -Width 4096).Trim()
        }
        else {
            $adminsText = (cmd /c 'net localgroup administrators' 2>$null) -join "`r`n"
        }
    }
    catch { $adminsText = $null }

    if ([string]::IsNullOrWhiteSpace($adminsText)) {
        $results.Add((New-EdcResult -Category 'Users' -CheckName 'Local Administrators group' -Status 'Not available' -Summary 'Unable to enumerate local Administrators group.' -Details 'Get-LocalGroupMember and net.exe fallback did not return data.' -RecommendedFix 'Run as Administrator; verify local accounts and policy.' )) | Out-Null
    }
    else {
        $results.Add((New-EdcResult -Category 'Users' -CheckName 'Local Administrators group' -Status 'OK' -Summary 'Administrators group listing succeeded.' -Details $adminsText -RecommendedFix '' )) | Out-Null
    }

    $localUsersText = $null
    $suspicious = New-Object System.Collections.Generic.List[string]
    try {
        if (Get-Command -Name Get-LocalUser -ErrorAction SilentlyContinue) {
            $users = Get-LocalUser -ErrorAction SilentlyContinue
            $localUsersText = ($users | Select-Object Name,Enabled,LastLogon,PasswordLastSet | Format-Table -AutoSize | Out-String -Width 4096).Trim()
            foreach ($u in $users) {
                if (-not $u.Enabled) { continue }
                if ($u.Name -in @('Administrator','DefaultAccount','Guest','WDAGUtilityAccount')) { continue }
                $suspicious.Add(("Enabled local account present: {0}" -f $u.Name)) | Out-Null
            }
        }
        else {
            $users = Get-CimInstance Win32_UserAccount -Filter "LocalAccount=True" -ErrorAction SilentlyContinue
            $localUsersText = ($users | Select-Object Name, Disabled, Lockout, PasswordChangeable, PasswordExpires | Format-Table -AutoSize | Out-String -Width 4096).Trim()
            foreach ($u in $users) {
                if ($u.Disabled) { continue }
                if ($u.Name -in @('Administrator','DefaultAccount','Guest','WDAGUtilityAccount')) { continue }
                $suspicious.Add(("Enabled local account present: {0}" -f $u.Name)) | Out-Null
            }
        }
    }
    catch { $localUsersText = $null }

    if ([string]::IsNullOrWhiteSpace($localUsersText)) {
        $results.Add((New-EdcResult -Category 'Users' -CheckName 'Local accounts' -Status 'Not available' -Summary 'Unable to enumerate local accounts.' -Details 'Local user enumeration cmdlets/WMI returned no data.' -RecommendedFix 'Run as Administrator; verify WMI service.' )) | Out-Null
    }
    else {
        $status = if ($suspicious.Count -gt 0) { 'Warning' } else { 'OK' }
        $summary = if ($status -eq 'Warning') { 'Review enabled local accounts (non-built-in).' } else { 'No obvious enabled non-built-in local accounts detected.' }
        $details = $localUsersText
        if ($suspicious.Count -gt 0) {
            $details = ($suspicious.ToArray() -join "`r`n") + "`r`n`r`n" + $localUsersText
        }
        $results.Add((New-EdcResult -Category 'Users' -CheckName 'Suspicious/stale local accounts' -Status $status -Summary $summary -Details $details -RecommendedFix 'Disable or remove unused local accounts and ensure strong passwords.')) | Out-Null
    }

    return $results.ToArray()
}

function Get-SecurityChecks {
    $results = New-Object System.Collections.Generic.List[object]

    $fwStatus = $null
    $fwDetails = ''
    try {
        if (Get-Command -Name Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
            $profiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select-Object Name, Enabled
            $fwDetails = ($profiles | Format-Table -AutoSize | Out-String -Width 4096).Trim()
            $enabledCount = @($profiles | Where-Object { $_.Enabled }).Count
            $fwStatus = if ($enabledCount -eq 0) { 'Failed' } elseif ($enabledCount -lt @($profiles).Count) { 'Warning' } else { 'OK' }
        }
        else {
            $out = (cmd /c 'netsh advfirewall show allprofiles state' 2>$null) -join "`r`n"
            if ($out -match 'State\s+ON') { $fwStatus = 'OK' } elseif ($out -match 'State\s+OFF') { $fwStatus = 'Warning' } else { $fwStatus = 'Not available' }
            $fwDetails = $out
        }
    }
    catch { $fwStatus = $null }

    if ($null -eq $fwStatus -or $fwStatus -eq 'Not available') {
        $results.Add((New-EdcResult -Category 'Security' -CheckName 'Firewall enabled' -Status 'Not available' -Summary 'Unable to determine firewall state.' -Details $fwDetails -RecommendedFix 'Run as Administrator and verify firewall configuration.' )) | Out-Null
    }
    else {
        $fix = ''
        $canFix = $false
        $fixAction = $null
        if ($fwStatus -ne 'OK' -and (Get-Command -Name Set-NetFirewallProfile -ErrorAction SilentlyContinue)) {
            $fix = 'Enable Windows Firewall profiles.'
            $canFix = $true
            $fixAction = { Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True -ErrorAction SilentlyContinue | Out-Null }
        }
        $results.Add((New-EdcResult -Category 'Security' -CheckName 'Firewall enabled' -Status $fwStatus -Summary $(if ($fwStatus -eq 'OK') { 'Firewall profiles appear enabled.' } elseif ($fwStatus -eq 'Failed') { 'All firewall profiles appear disabled.' } else { 'One or more firewall profiles are disabled.' }) -Details $fwDetails -RecommendedFix $fix -CanAutoFix:$canFix -FixAction $fixAction -FixRequiresAdmin:$true -FixDisruptive:$true)) | Out-Null
    }

    try {
        if (Get-Command -Name Get-MpComputerStatus -ErrorAction SilentlyContinue) {
            $mp = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if ($mp) {
                $rtp = [bool]$mp.RealTimeProtectionEnabled
                $am = [bool]$mp.AntivirusEnabled
                $status = if ($am -and $rtp) { 'OK' } else { 'Warning' }
                $results.Add((New-EdcResult -Category 'Security' -CheckName 'Defender status' -Status $status -Summary ("AntivirusEnabled={0}, RealTimeProtectionEnabled={1}" -f $am, $rtp) -Details ($mp | Format-List * | Out-String -Width 4096).Trim() -RecommendedFix 'If managed by policy, follow org guidance; otherwise ensure Defender is enabled and up to date.' )) | Out-Null
            }
            else {
                $results.Add((New-EdcResult -Category 'Security' -CheckName 'Defender status' -Status 'Not available' -Summary 'Get-MpComputerStatus returned no data.' -Details '' -RecommendedFix '' )) | Out-Null
            }
        }
        else {
            $svc = Get-Service -Name 'WinDefend' -ErrorAction SilentlyContinue
            if ($svc) {
                $status = if ($svc.Status -eq 'Running') { 'OK' } else { 'Warning' }
                $results.Add((New-EdcResult -Category 'Security' -CheckName 'Defender status' -Status $status -Summary ("Service WinDefend is {0}." -f $svc.Status) -Details ($svc | Format-List * | Out-String -Width 4096).Trim() -RecommendedFix 'If Defender is expected, ensure the service is running and not disabled by policy.' )) | Out-Null
            }
            else {
                $results.Add((New-EdcResult -Category 'Security' -CheckName 'Defender status' -Status 'Not available' -Summary 'Defender cmdlets/service not found.' -Details '' -RecommendedFix 'This system may use a different AV solution.' )) | Out-Null
            }
        }
    }
    catch {
        $results.Add((New-EdcResult -Category 'Security' -CheckName 'Defender status' -Status 'Not available' -Summary 'Unable to query Defender status.' -Details $_.Exception.Message -RecommendedFix '' )) | Out-Null
    }

    try {
        if (Get-Command -Name Get-BitLockerVolume -ErrorAction SilentlyContinue) {
            $vols = Get-BitLockerVolume -ErrorAction SilentlyContinue | Select-Object MountPoint, VolumeStatus, ProtectionStatus, EncryptionPercentage
            if ($vols) {
                $details = ($vols | Format-Table -AutoSize | Out-String -Width 4096).Trim()
                $results.Add((New-EdcResult -Category 'Security' -CheckName 'BitLocker status' -Status 'OK' -Summary 'BitLocker query succeeded.' -Details $details -RecommendedFix 'Follow org policy for encryption requirements.' )) | Out-Null
            }
            else {
                $results.Add((New-EdcResult -Category 'Security' -CheckName 'BitLocker status' -Status 'Not available' -Summary 'Get-BitLockerVolume returned no data.' -Details '' -RecommendedFix '' )) | Out-Null
            }
        }
        else {
            $results.Add((New-EdcResult -Category 'Security' -CheckName 'BitLocker status' -Status 'Not available' -Summary 'BitLocker cmdlets not available.' -Details 'Get-BitLockerVolume is not present on this endpoint.' -RecommendedFix '' )) | Out-Null
        }
    }
    catch {
        $results.Add((New-EdcResult -Category 'Security' -CheckName 'BitLocker status' -Status 'Not available' -Summary 'Unable to query BitLocker.' -Details $_.Exception.Message -RecommendedFix '' )) | Out-Null
    }

    return $results.ToArray()
}

function Get-ServicesChecks {
    $results = New-Object System.Collections.Generic.List[object]
    $critical = @('Spooler','BITS','wuauserv')

    foreach ($name in $critical) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if (-not $svc) {
            $results.Add((New-EdcResult -Category 'Services' -CheckName ("Service: {0}" -f $name) -Status 'Not available' -Summary 'Service not found.' -Details '' -RecommendedFix '' )) | Out-Null
            continue
        }

        $status = if ($svc.Status -eq 'Running') { 'OK' } else { 'Warning' }
        $fix = ''
        $canFix = $false
        $fixAction = $null
        if ($svc.Status -ne 'Running') {
            $svcName = $name
            $fix = "Restart/start the $name service."
            $canFix = $true
            $fixAction = { Restart-Service -Name $svcName -Force -ErrorAction SilentlyContinue }.GetNewClosure()
        }
        $details = ($svc | Format-List * | Out-String -Width 4096).Trim()
        $results.Add((New-EdcResult -Category 'Services' -CheckName ("Service: {0}" -f $name) -Status $status -Summary ("{0} is {1}." -f $name, $svc.Status) -Details $details -RecommendedFix $fix -CanAutoFix:$canFix -FixAction $fixAction -FixRequiresAdmin:$true -FixDisruptive:$true)) | Out-Null
    }

    return $results.ToArray()
}

function Get-FileSystemChecks {
    $results = New-Object System.Collections.Generic.List[object]

    $drives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
    if (-not $drives) {
        $results.Add((New-EdcResult -Category 'File System' -CheckName 'Drive space' -Status 'Not available' -Summary 'No fixed disks returned.' -Details '' -RecommendedFix '' )) | Out-Null
    }
    else {
        $lines = New-Object System.Collections.Generic.List[string]
        $worst = 'OK'
        foreach ($d in $drives) {
            if ($d.Size -eq $null -or $d.FreeSpace -eq $null) { continue }
            $total = [double]$d.Size
            if ($total -le 0) { continue }
            $free = [double]$d.FreeSpace
            $pctFree = [math]::Round(($free / $total) * 100, 2)
            $lines.Add(("{0}: {1}% free ({2} GB free)" -f $d.DeviceID, $pctFree, [math]::Round(($free/1GB),2))) | Out-Null
            if ($pctFree -lt 10) { $worst = 'Failed' }
            elseif ($pctFree -lt 15 -and $worst -ne 'Failed') { $worst = 'Warning' }
        }
        $results.Add((New-EdcResult -Category 'File System' -CheckName 'Drive space' -Status $worst -Summary $(if ($lines.Count -gt 0) { $lines[0] } else { 'Drive space counters unavailable.' }) -Details (($lines.ToArray() -join "`r`n")) -RecommendedFix $(if ($worst -eq 'OK') { '' } else { 'Free space by removing temporary files and unused data.' }) )) | Out-Null
    }

    $temp = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($temp) -or -not (Test-Path -Path $temp)) {
        $results.Add((New-EdcResult -Category 'File System' -CheckName 'Temp accumulation' -Status 'Not available' -Summary 'TEMP folder not found.' -Details ("TEMP={0}" -f $temp) -RecommendedFix '' )) | Out-Null
    }
    else {
        $count = $null
        try { $count = (Get-ChildItem -Path $temp -Force -ErrorAction SilentlyContinue | Measure-Object).Count } catch { $count = $null }
        if ($null -eq $count) {
            $results.Add((New-EdcResult -Category 'File System' -CheckName 'Temp accumulation' -Status 'Not available' -Summary 'Unable to enumerate TEMP folder.' -Details $temp -RecommendedFix '' )) | Out-Null
        }
        else {
            $status = if ($count -ge 15000) { 'Failed' } elseif ($count -ge 5000) { 'Warning' } else { 'OK' }
            $fix = ''
            $canFix = $false
            $fixAction = $null
            if ($status -ne 'OK') {
                $fix = 'Delete contents of the current user TEMP folder.'
                $canFix = $true
                $fixAction = {
                    $tempPath = $env:TEMP
                    if ([string]::IsNullOrWhiteSpace($tempPath) -or -not (Test-Path -Path $tempPath)) { return }
                    Get-ChildItem -Path $tempPath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            $results.Add((New-EdcResult -Category 'File System' -CheckName 'Temp accumulation' -Status $status -Summary ("TEMP contains ~{0} top-level items." -f $count) -Details $temp -RecommendedFix $fix -CanAutoFix:$canFix -FixAction $fixAction -FixRequiresAdmin:$false -FixDisruptive:$true)) | Out-Null
        }
    }

    return $results.ToArray()
}

function Get-ToolsChecks {
    $results = New-Object System.Collections.Generic.List[object]

    $toolSpecs = @(
        @{ Name = 'System File Checker (sfc)'; Command = 'sfc'; Hint = 'Use `sfc /scannow` in an elevated shell for system file repair.' }
        @{ Name = 'DISM'; Command = 'dism'; Hint = 'Use DISM for image health checks and driver export.' }
        @{ Name = 'CHKDSK'; Command = 'chkdsk'; Hint = 'Use chkdsk for disk integrity checks during maintenance windows.' }
        @{ Name = 'winget'; Command = 'winget'; Hint = 'Use winget for package inventory and upgrades where approved.' }
        @{ Name = 'PowerCfg'; Command = 'powercfg'; Hint = 'Use powercfg for sleep/battery diagnostics.' }
    )

    foreach ($spec in $toolSpecs) {
        $cmd = Get-Command -Name $spec.Command -ErrorAction SilentlyContinue
        if ($cmd) {
            $summary = "{0} available at {1}" -f $spec.Command, $cmd.Source
            $results.Add((New-EdcResult -Category 'Tools / Utilities' -CheckName $spec.Name -Status 'OK' -Summary $summary -Details $spec.Hint -ResultGroup 'Tool Availability')) | Out-Null
        }
        else {
            $summary = "{0} command was not found." -f $spec.Command
            $results.Add((New-EdcResult -Category 'Tools / Utilities' -CheckName $spec.Name -Status 'Not available' -Summary $summary -Details $spec.Hint -ResultGroup 'Tool Availability')) | Out-Null
        }
    }

    $adminTools = @(
        @{ Name = 'Event Viewer'; Launcher = 'eventvwr.msc' }
        @{ Name = 'Device Manager'; Launcher = 'devmgmt.msc' }
        @{ Name = 'Services Console'; Launcher = 'services.msc' }
        @{ Name = 'Computer Management'; Launcher = 'compmgmt.msc' }
    )
    $details = ($adminTools | ForEach-Object { "{0} -> {1}" -f $_.Name, $_.Launcher }) -join "`r`n"
    $results.Add((New-EdcResult -Category 'Tools / Utilities' -CheckName 'Windows admin console launchers' -Status 'OK' -Summary 'Quick-launch entries are ready from the Actions panel.' -Details $details -ResultGroup 'Quick Launch')) | Out-Null

    return $results.ToArray()
}

$script:CategoryToolkitHandlers = @{
    'System'      = @('Get-OSInfo','Get-SystemUptime','Get-DiskInfo','Get-RAMInfo')
    'Network'     = @('Get-IPConfiguration','Get-AdapterSummary','Get-ListeningPorts')
    'File System' = @('Get-DriveFreeSpaceSummary','Get-RecentFilesListing')
    'Users'       = @('Get-LocalUsersList','Get-AdministratorsGroupMembers','Get-LastLogonInfo')
    'Security'    = @('Get-FirewallStatus','Get-DefenderStatus','Get-BitLockerStatus')
    'Services'    = @('Get-ServicesList')
    'Tools / Utilities' = @()
}

function Invoke-CategoryScan {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$PingTarget,
        [string]$DnsTarget = 'www.microsoft.com'
    )

    $out = New-Object System.Collections.Generic.List[object]

    if ($script:CategoryToolkitHandlers.ContainsKey($Category)) {
        $handlers = $script:CategoryToolkitHandlers[$Category]
        if ($Category -eq 'Services') { $handlers = @() }
        if ($Category -eq 'File System') { $handlers = @('Get-DriveFreeSpaceSummary') }
        if ($handlers.Count -gt 0) {
            $out.Add((Invoke-ToolkitHandlersAsInventory -Category $Category -Handlers $handlers)) | Out-Null
        }
    }

    switch ($Category) {
        'System'      { (Get-SystemChecks) | ForEach-Object { $out.Add($_) | Out-Null } }
        'Network'     { (Get-NetworkChecks -PingTarget $PingTarget -DnsTarget $DnsTarget) | ForEach-Object { $out.Add($_) | Out-Null } }
        'File System' { (Get-FileSystemChecks) | ForEach-Object { $out.Add($_) | Out-Null } }
        'Users'       { (Get-UsersChecks) | ForEach-Object { $out.Add($_) | Out-Null } }
        'Security'    { (Get-SecurityChecks) | ForEach-Object { $out.Add($_) | Out-Null } }
        'Services'    { (Get-ServicesChecks) | ForEach-Object { $out.Add($_) | Out-Null } }
        'Tools / Utilities' { (Get-ToolsChecks) | ForEach-Object { $out.Add($_) | Out-Null } }
        default { }
    }

    return $out.ToArray()
}

# -----------------------------
# GUI construction
# -----------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'EDC Toolkit Beta'
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $script:Palette.FormBack
$form.KeyPreview = $true
$form.MinimumSize = New-Object System.Drawing.Size(1180, 820)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = 'Ready.'
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

$logoBox = New-Object System.Windows.Forms.PictureBox
$logoBox.Location = New-Object System.Drawing.Point(20, 15)
$logoBox.Size = New-Object System.Drawing.Size(80, 80)
$logoBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$logoPath = Join-Path $PSScriptRoot 'logo.png'
if (Test-Path $logoPath) {
    try { $logoBox.Image = [System.Drawing.Image]::FromFile($logoPath) } catch { }
}
$form.Controls.Add($logoBox)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'EDC Toolkit'
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(120, 20)
$titleLabel.Size = New-Object System.Drawing.Size(420, 40)
$form.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Name = 'SubtitleLabel'
$subtitleLabel.Text = 'Beta GUI wrapper for endpoint triage'
$subtitleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Regular)
$subtitleLabel.ForeColor = $script:Palette.MutedText
$subtitleLabel.Location = New-Object System.Drawing.Point(122, 60)
$subtitleLabel.Size = New-Object System.Drawing.Size(460, 22)
$form.Controls.Add($subtitleLabel)

$separator = New-Object System.Windows.Forms.Panel
$separator.Name = 'SeparatorPanel'
$separator.Location = New-Object System.Drawing.Point(20, 105)
$separator.Size = New-Object System.Drawing.Size(1120, 2)
$separator.BackColor = $script:Palette.Border
$form.Controls.Add($separator)

$summaryGroup = New-Object System.Windows.Forms.GroupBox
$summaryGroup.Text = 'Overview'
$summaryGroup.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$summaryGroup.Location = New-Object System.Drawing.Point(20, 120)
$summaryGroup.Size = New-Object System.Drawing.Size(500, 82)
$form.Controls.Add($summaryGroup)

$lblQuickSummary = New-Object System.Windows.Forms.Label
$lblQuickSummary.Text = 'Summary: Passed=0  Warnings=0  Failed=0  N/A=0'
$lblQuickSummary.Dock = 'Fill'
$lblQuickSummary.TextAlign = 'MiddleLeft'
$lblQuickSummary.Padding = New-Object System.Windows.Forms.Padding(12, 0, 8, 0)
$lblQuickSummary.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$summaryGroup.Controls.Add($lblQuickSummary)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 20000
$toolTip.InitialDelay = 300
$toolTip.ReshowDelay = 100
$toolTip.ShowAlways = $true

$actionsGroup = New-Object System.Windows.Forms.GroupBox
$actionsGroup.Text = 'Actions'
$actionsGroup.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$actionsGroup.Location = New-Object System.Drawing.Point(540, 120)
$actionsGroup.Size = New-Object System.Drawing.Size(600, 186)
$form.Controls.Add($actionsGroup)

$actionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionsLayout.Dock = 'Fill'
$actionsLayout.Padding = '14,18,14,14'
$actionsLayout.ColumnCount = 4
$actionsLayout.RowCount = 3
$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$actionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54))) | Out-Null
$actionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54))) | Out-Null
$actionsGroup.Controls.Add($actionsLayout)

$pingPanel = New-Object System.Windows.Forms.Panel
$pingPanel.Dock = 'Fill'

$lblPing = New-Object System.Windows.Forms.Label
$lblPing.Text = 'Ping Target:'
$lblPing.AutoSize = $true
$lblPing.Location = New-Object System.Drawing.Point(2, 6)
$pingPanel.Controls.Add($lblPing)

$txtPing = New-Object System.Windows.Forms.TextBox
$txtPing.Text = $script:PingTarget
$txtPing.Location = New-Object System.Drawing.Point(82, 2)
$txtPing.Size = New-Object System.Drawing.Size(160, 23)
$pingPanel.Controls.Add($txtPing)

$lblDns = New-Object System.Windows.Forms.Label
$lblDns.Text = 'DNS Name:'
$lblDns.AutoSize = $true
$lblDns.Location = New-Object System.Drawing.Point(258, 6)
$pingPanel.Controls.Add($lblDns)

$txtDns = New-Object System.Windows.Forms.TextBox
$txtDns.Text = $script:DnsTarget
$txtDns.Location = New-Object System.Drawing.Point(332, 2)
$txtDns.Size = New-Object System.Drawing.Size(238, 23)
$pingPanel.Controls.Add($txtDns)

$null = $actionsLayout.Controls.Add($pingPanel, 0, 0)
$actionsLayout.SetColumnSpan($pingPanel, 4)

function New-ActionButton {
    param([Parameter(Mandatory)][string]$Text)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Dock = 'Fill'
    $btn.Margin = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
    $btn.MinimumSize = New-Object System.Drawing.Size(128, 42)
    $btn.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    return $btn
}

$btnFullScan = New-ActionButton -Text 'Full Scan'
$btnTrouble  = New-ActionButton -Text 'Troubleshoot'
$btnExport   = New-ActionButton -Text 'Export Report'
$btnFix      = New-ActionButton -Text 'Apply Fixes'
$btnOpenRpt  = New-ActionButton -Text 'Open Reports'
$btnEventVwr = New-ActionButton -Text 'Event Viewer'
$btnDevMgr   = New-ActionButton -Text 'Device Manager'
$btnHelp     = New-ActionButton -Text 'Help'

$null = $actionsLayout.Controls.Add($btnFullScan, 0, 1)
$null = $actionsLayout.Controls.Add($btnTrouble, 1, 1)
$null = $actionsLayout.Controls.Add($btnExport, 2, 1)
$null = $actionsLayout.Controls.Add($btnHelp, 3, 1)
$null = $actionsLayout.Controls.Add($btnFix, 0, 2)
$null = $actionsLayout.Controls.Add($btnOpenRpt, 1, 2)
$null = $actionsLayout.Controls.Add($btnEventVwr, 2, 2)
$null = $actionsLayout.Controls.Add($btnDevMgr, 3, 2)

$toolTip.SetToolTip($btnFullScan, 'Runs a full scan across all categories and captures selected outputs from the existing toolkit reports.')
$toolTip.SetToolTip($btnTrouble, 'Runs conservative health checks across all categories (faster, less report capture).')
$toolTip.SetToolTip($btnFix, 'Applies safe automated fixes for items marked Warning/Failed that support auto-fix. Disruptive actions require confirmation.')
$toolTip.SetToolTip($btnExport, 'Exports the current results to a single text report (and also writes a copy into the toolkit report session when available).')
$toolTip.SetToolTip($btnOpenRpt, 'Opens the toolkit report folder for this GUI session.')
$toolTip.SetToolTip($btnEventVwr, 'Launches Event Viewer.')
$toolTip.SetToolTip($btnDevMgr, 'Launches Device Manager.')
$toolTip.SetToolTip($btnHelp, 'Explains what the buttons and tabs do.')
$toolTip.SetToolTip($txtPing, 'IPv4/host used by the ping check in the Network tab.')
$toolTip.SetToolTip($txtDns, 'DNS name used by the Network DNS resolution check.')

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(20, 326)
$tabs.Size = New-Object System.Drawing.Size(1120, 468)
$form.Controls.Add($tabs)

$imgList = New-StatusImageList

$script:CategoryViews = @{}
$script:ScanCategories = @('System','Network','File System','Users','Security','Services','Tools / Utilities')
$script:AllCategories = @($script:ScanCategories + @('Reports / Logs'))
$script:CategoryDisplay = @{
    'System'            = 'System'
    'Network'           = 'Network'
    'File System'       = 'File System'
    'Users'             = 'Users'
    'Security'          = 'Security'
    'Services'          = 'Services'
    'Tools / Utilities' = 'Tools / Utilities'
    'Reports / Logs'    = 'Reports / Logs'
}

function Get-CategoryFilters {
    param([Parameter(Mandatory)][string]$Category)

    switch ($Category) {
        'Reports / Logs' {
            return @(
                [pscustomobject]@{ Text='[ALL] All result groups'; Value='*' }
                [pscustomobject]@{ Text='[GUI] GUI session'; Value='GUI Session' }
                [pscustomobject]@{ Text='[RPT] Report files'; Value='Report Files' }
                [pscustomobject]@{ Text='[TK] Toolkit output'; Value='Toolkit Output' }
            )
        }
        'Tools / Utilities' {
            return @(
                [pscustomobject]@{ Text='[ALL] All result groups'; Value='*' }
                [pscustomobject]@{ Text='[CHK] Tool availability'; Value='Tool Availability' }
                [pscustomobject]@{ Text='[RUN] Quick launch'; Value='Quick Launch' }
            )
        }
        default {
            return @(
                [pscustomobject]@{ Text='[ALL] All result groups'; Value='*' }
                [pscustomobject]@{ Text='[HC] Health checks'; Value='Health Checks' }
                [pscustomobject]@{ Text='[TK] Toolkit output'; Value='Toolkit Output' }
            )
        }
    }
}

function Add-CategoryTab {
    param([Parameter(Mandatory)][string]$Category)

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = if ($script:CategoryDisplay.ContainsKey($Category)) { $script:CategoryDisplay[$Category] } else { $Category }
    $tab.BackColor = $script:Palette.Surface

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = 'Fill'
    $split.Orientation = 'Vertical'
    $split.SplitterDistance = 420
    $split.BackColor = $script:Palette.Surface

    $leftPanel = New-Object System.Windows.Forms.Panel
    $leftPanel.Dock = 'Fill'

    $leftTop = New-Object System.Windows.Forms.Panel
    $leftTop.Dock = 'Top'
    $leftTop.Height = 34

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Scan $Category"
    $btnScan.Location = New-Object System.Drawing.Point(0, 2)
    $btnScan.Size = New-Object System.Drawing.Size(130, 30)
    $btnScan.Height = 30

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = 'Result Group:'
    $lblFilter.AutoSize = $true
    $lblFilter.Location = New-Object System.Drawing.Point(140, 9)

    $cmbFilter = New-Object System.Windows.Forms.ComboBox
    $cmbFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbFilter.Location = New-Object System.Drawing.Point(225, 5)
    $cmbFilter.Size = New-Object System.Drawing.Size(190, 24)

    foreach ($f in (Get-CategoryFilters -Category $Category)) {
        $null = $cmbFilter.Items.Add($f)
    }
    $cmbFilter.DisplayMember = 'Text'
    $cmbFilter.ValueMember = 'Value'
    if ($cmbFilter.Items.Count -gt 0) { $cmbFilter.SelectedIndex = 0 }

    $leftTop.Controls.Add($btnScan)
    $leftTop.Controls.Add($lblFilter)
    $leftTop.Controls.Add($cmbFilter)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = 'Fill'
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    $lv.HideSelection = $false
    $lv.SmallImageList = $imgList
    $null = $lv.Columns.Add('Status', 90)
    $null = $lv.Columns.Add('Check', 180)
    $null = $lv.Columns.Add('Summary', 420)

    $leftPanel.Controls.Add($lv)
    $leftPanel.Controls.Add($leftTop)

    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Dock = 'Fill'

    $catSummary = New-Object System.Windows.Forms.Label
    $catSummary.Text = 'Summary: (no results yet)'
    $catSummary.Dock = 'Top'
    $catSummary.Height = 26
    $catSummary.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

    $details = New-Object System.Windows.Forms.RichTextBox
    $details.Dock = 'Fill'
    $details.Font = New-Object System.Drawing.Font('Consolas', 9)
    $details.ReadOnly = $true

    $rightPanel.Controls.Add($details)
    $rightPanel.Controls.Add($catSummary)

    $split.Panel1.Controls.Add($leftPanel)
    $split.Panel2.Controls.Add($rightPanel)

    $tab.Controls.Add($split)
    $tabs.TabPages.Add($tab) | Out-Null

    $script:CategoryViews[$Category] = @{
        Tab        = $tab
        ScanButton = $btnScan
        FilterCombo= $cmbFilter
        ListView   = $lv
        Details    = $details
        Summary    = $catSummary
    }
}

$script:AllCategories | ForEach-Object { Add-CategoryTab -Category $_ }
Apply-ControlTheme -Control $form -Palette $script:Palette

function Update-QuickSummary {
    $all = Get-AllResults
    $passed = @($all | Where-Object { $_.Status -eq 'OK' }).Count
    $warn = @($all | Where-Object { $_.Status -eq 'Warning' }).Count
    $fail = @($all | Where-Object { $_.Status -eq 'Failed' }).Count
    $na = @($all | Where-Object { $_.Status -eq 'Not available' }).Count
    $lblQuickSummary.Text = "Summary: Passed=$passed  Warnings=$warn  Failed=$fail  N/A=$na"
}

function Refresh-CategoryView {
    param([Parameter(Mandatory)][string]$Category)
    if (-not $script:CategoryViews.ContainsKey($Category)) { return }

    $view = $script:CategoryViews[$Category]
    $selectedGroup = '*'
    if ($view.FilterCombo -and $view.FilterCombo.SelectedItem) {
        $selectedGroup = [string]$view.FilterCombo.SelectedItem.Value
        if ([string]::IsNullOrWhiteSpace($selectedGroup)) { $selectedGroup = '*' }
    }

    $lv = $view.ListView
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        $allCategoryItems = @(Get-CategoryResults -Category $Category)
        $items = @(
            if ($selectedGroup -eq '*') {
            $allCategoryItems
        }
        else {
                @($allCategoryItems | Where-Object { $_.ResultGroup -eq $selectedGroup })
            }
        )

        foreach ($r in $items) {
            $item = New-Object System.Windows.Forms.ListViewItem
            $item.Text = $r.Status
            $item.ImageKey = Get-StatusImageKey -Status $r.Status
            $null = $item.SubItems.Add($r.CheckName)
            $null = $item.SubItems.Add($r.Summary)
            $item.Tag = $r
            $lv.Items.Add($item) | Out-Null
        }

        $passed = @($allCategoryItems | Where-Object { $_.Status -eq 'OK' }).Count
        $warn = @($allCategoryItems | Where-Object { $_.Status -eq 'Warning' }).Count
        $fail = @($allCategoryItems | Where-Object { $_.Status -eq 'Failed' }).Count
        $na = @($allCategoryItems | Where-Object { $_.Status -eq 'Not available' }).Count
        $shown = @($items).Count
        $view.Summary.Text = "Showing $shown item(s) in current filter."
    }
    finally {
        $lv.EndUpdate()
    }
}

function Refresh-AllViews {
    foreach ($cat in $script:CategoryViews.Keys) { Refresh-CategoryView -Category $cat }
    Update-QuickSummary
}

function Show-ResultDetails {
    param(
        [Parameter(Mandatory)][string]$Category,
        [AllowNull()][object]$Result
    )
    if (-not $script:CategoryViews.ContainsKey($Category)) { return }
    $box = $script:CategoryViews[$Category].Details
    if ($null -eq $Result) { $box.Text = ''; return }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(("Category: {0}" -f $Result.Category)) | Out-Null
    $lines.Add(("Group:    {0}" -f $Result.ResultGroup)) | Out-Null
    $lines.Add(("Check:    {0}" -f $Result.CheckName)) | Out-Null
    $lines.Add(("Status:   {0}" -f $Result.Status)) | Out-Null
    $lines.Add(("When:     {0}" -f $Result.CollectedAt.ToString('yyyy-MM-dd HH:mm:ss'))) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($Result.ReportPath)) {
        $lines.Add(("File:     {0}" -f $Result.ReportPath)) | Out-Null
    }
    $lines.Add('') | Out-Null
    $lines.Add('Summary:') | Out-Null
    $lines.Add($Result.Summary) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Details:') | Out-Null
    $lines.Add($Result.Details) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($Result.RecommendedFix)) {
        $lines.Add('') | Out-Null
        $lines.Add('Recommended Fix:') | Out-Null
        $lines.Add($Result.RecommendedFix) | Out-Null
        $lines.Add(("Automatable: {0}" -f $(if ($Result.CanAutoFix) { 'Yes' } else { 'No' }))) | Out-Null
        if ($Result.FixRequiresAdmin) { $lines.Add('Requires admin: Yes') | Out-Null }
        if ($Result.FixDisruptive) { $lines.Add('Disruptive: Yes (confirmation required)') | Out-Null }
    }
    $box.Text = ($lines.ToArray() -join "`r`n")
}

$script:RefreshCategoryViewAction = ${function:Refresh-CategoryView}
$script:ShowResultDetailsAction = ${function:Show-ResultDetails}

foreach ($cat in $script:CategoryViews.Keys) {
    $catLocal = $cat
    $view = $script:CategoryViews[$catLocal]
    if ($view.FilterCombo) {
        $view.FilterCombo.add_SelectedIndexChanged({
            & $script:RefreshCategoryViewAction -Category $catLocal
            & $script:ShowResultDetailsAction -Category $catLocal -Result $null
        }.GetNewClosure())
    }
    $view.ListView.add_SelectedIndexChanged({
        $lv = $this
        $picked = $null
        if ($lv.SelectedItems.Count -gt 0) { $picked = $lv.SelectedItems[0].Tag }
        & $script:ShowResultDetailsAction -Category $catLocal -Result $picked
    }.GetNewClosure())
}

# Double-click to open report files from Reports/Logs.
if ($script:CategoryViews.ContainsKey('Reports / Logs')) {
    $repView = $script:CategoryViews['Reports / Logs']
    $repView.ListView.add_DoubleClick({
        try {
            if ($this.SelectedItems.Count -lt 1) { return }
            $r = $this.SelectedItems[0].Tag
            if ($null -eq $r) { return }
            if (-not [string]::IsNullOrWhiteSpace($r.ReportPath) -and (Test-Path -Path $r.ReportPath)) {
                Start-Process -FilePath $r.ReportPath | Out-Null
            }
        }
        catch { }
    })
}

# -----------------------------
# Scanning (simple + reliable)
# -----------------------------

$script:CurrentWorkLabel = ''
$script:IsBusy = $false
$script:GuiLog = New-Object System.Collections.Generic.List[string]

function Add-GuiLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $script:GuiLog.Add($line) | Out-Null

    $text = ($script:GuiLog.ToArray() -join "`r`n")
    $rep = New-EdcResult -Category 'Reports / Logs' -CheckName 'GUI Log' -Status 'OK' -Summary ("{0} log lines" -f $script:GuiLog.Count) -Details $text -ResultGroup 'GUI Session'
    Merge-Results -Results @($rep)
    Refresh-CategoryView -Category 'Reports / Logs'
}

function Set-UiBusy {
    param([bool]$Busy,[string]$Message)
    $btnFullScan.Enabled = -not $Busy
    $btnTrouble.Enabled = -not $Busy
    $btnFix.Enabled = -not $Busy
    $btnExport.Enabled = -not $Busy
    $btnOpenRpt.Enabled = -not $Busy
    $btnEventVwr.Enabled = -not $Busy
    $btnDevMgr.Enabled = -not $Busy
    $btnHelp.Enabled = -not $Busy
    $txtPing.Enabled = -not $Busy
    $txtDns.Enabled = -not $Busy
    foreach ($cat in $script:CategoryViews.Keys) { $script:CategoryViews[$cat].ScanButton.Enabled = -not $Busy }
    $statusLabel.Text = $Message
}

function Start-Scan {
    param(
        [Parameter(Mandatory)][string]$Mode,
        [Parameter(Mandatory)][string[]]$Categories
    )
    if ($script:IsBusy) { return }
    $script:IsBusy = $true
    try {
        $script:PingTarget = $txtPing.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($script:PingTarget)) {
            $script:PingTarget = '8.8.8.8'
            $txtPing.Text = $script:PingTarget
        }
        $script:DnsTarget = $txtDns.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($script:DnsTarget)) {
            $script:DnsTarget = 'www.microsoft.com'
            $txtDns.Text = $script:DnsTarget
        }
        $script:CurrentWorkLabel = $Mode
        Add-GuiLog -Level INFO -Message ("Start: {0}" -f $Mode)
        Set-UiBusy -Busy:$true -Message ("Running: {0}" -f $Mode)

        $i = 0
        foreach ($c in $Categories) {
            $i++
            $statusLabel.Text = ("{0}: {1} ({2}/{3})" -f $Mode, $c, $i, $Categories.Count)
            [System.Windows.Forms.Application]::DoEvents()

            $results = @(Invoke-CategoryScan -Category $c -PingTarget $script:PingTarget -DnsTarget $script:DnsTarget)
            if ($results.Count -gt 0) { Merge-Results -Results $results }
            Refresh-CategoryView -Category $c
            Update-QuickSummary
            [System.Windows.Forms.Application]::DoEvents()
        }

        Sync-ReportFilesToResults
        Refresh-CategoryView -Category 'Reports / Logs'
        Update-QuickSummary

        $statusLabel.Text = "Completed: $script:CurrentWorkLabel"
        Add-GuiLog -Level INFO -Message ("Completed: {0}" -f $Mode)
    }
    catch {
        $msg = $_.Exception.Message
        Add-GuiLog -Level ERROR -Message ("{0} failed: {1}" -f $Mode, $msg)
        [System.Windows.Forms.MessageBox]::Show($msg, 'EDC Toolkit Beta', 'OK', 'Error') | Out-Null
    }
    finally {
        Set-UiBusy -Busy:$false -Message 'Ready.'
        $script:IsBusy = $false
    }
}

# -----------------------------
# Fix application + export
# -----------------------------

function Apply-RecommendFixes {
    if ($script:IsBusy) { return }

    $targets = @(Get-AllResults | Where-Object { ($_.Status -in @('Warning','Failed')) -and $_.CanAutoFix -and $_.FixAction })
    if (-not $targets -or $targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No recommended automated fixes are currently available.', 'EDC Toolkit Beta', 'OK', 'Information') | Out-Null
        return
    }

    $admin = Get-IsAdmin
    $needAdmin = @($targets | Where-Object { $_.FixRequiresAdmin }).Count -gt 0
    if ($needAdmin -and -not $admin) {
        [System.Windows.Forms.MessageBox]::Show('Some fixes may require Administrator. Re-run as Administrator if fixes fail.', 'EDC Toolkit Beta', 'OK', 'Warning') | Out-Null
    }

    $list = ($targets | Select-Object Category,CheckName,RecommendedFix | ForEach-Object { "- [{0}] {1}: {2}" -f $_.Category, $_.CheckName, $_.RecommendedFix }) -join "`r`n"
    if (-not (Confirm-GuiAction -Message ("Apply recommended fixes?`r`n`r`n{0}" -f $list))) { return }

    foreach ($t in $targets) {
        if ($t.FixDisruptive) {
            if (-not (Confirm-GuiAction -Message ("This fix can be disruptive:`r`n[{0}] {1}`r`n`r`n{2}`r`n`r`nProceed?" -f $t.Category, $t.CheckName, $t.RecommendedFix))) {
                continue
            }
        }
        try {
            $statusLabel.Text = ("Applying fix: [{0}] {1}" -f $t.Category, $t.CheckName)
            & $t.FixAction
        }
        catch { }
    }

    Start-Scan -Mode 'Re-Scan' -Categories $script:ScanCategories
}

function Export-ResultsReport {
    if ($script:IsBusy) { return }
    $all = @(Get-AllResults)
    if (-not $all -or $all.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No results to export yet. Run a scan first.', 'EDC Toolkit Beta', 'OK', 'Information') | Out-Null
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = 'Export EDC Toolkit Report'
    $dlg.Filter = 'Text Report (*.txt)|*.txt|All Files (*.*)|*.*'
    $dlg.FileName = ('EDCtoolkit_Report_{0}.txt' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $passed = @($all | Where-Object { $_.Status -eq 'OK' }).Count
    $warn = @($all | Where-Object { $_.Status -eq 'Warning' }).Count
    $fail = @($all | Where-Object { $_.Status -eq 'Failed' }).Count
    $na = @($all | Where-Object { $_.Status -eq 'Not available' }).Count

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("EDC Toolkit Beta Report")
    [void]$sb.AppendLine(("Generated: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$sb.AppendLine(("Host: {0}" -f $env:COMPUTERNAME))
    [void]$sb.AppendLine(("User: {0}" -f $env:USERNAME))
    [void]$sb.AppendLine(("Admin: {0}" -f $(if (Get-IsAdmin) { 'Yes' } else { 'No' })))
    [void]$sb.AppendLine(('=' * 72))
    [void]$sb.AppendLine(("Quick Summary: Passed={0}  Warnings={1}  Failed={2}  N/A={3}" -f $passed,$warn,$fail,$na))
    [void]$sb.AppendLine(('=' * 72))

    foreach ($cat in $script:AllCategories) {
        $items = Get-CategoryResults -Category $cat
        if (-not $items -or $items.Count -eq 0) { continue }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(("[{0}]" -f $cat))
        foreach ($r in $items) {
            [void]$sb.AppendLine(("- {0} | {1} | {2} | Group={3}" -f $r.Status, $r.CheckName, $r.Summary, $r.ResultGroup))
            if (-not [string]::IsNullOrWhiteSpace($r.RecommendedFix) -and $r.Status -ne 'OK') {
                [void]$sb.AppendLine(("    Fix: {0}" -f $r.RecommendedFix))
                [void]$sb.AppendLine(("    Auto: {0}" -f $(if ($r.CanAutoFix) { 'Yes' } else { 'No' })))
            }
        }
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('=' * 72))
    [void]$sb.AppendLine('Detailed Sections')
    [void]$sb.AppendLine(('=' * 72))

    foreach ($r in $all) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(("[{0}] {1} ({2}) [Group: {3}]" -f $r.Category, $r.CheckName, $r.Status, $r.ResultGroup))
        [void]$sb.AppendLine($r.Details)
        if (-not [string]::IsNullOrWhiteSpace($r.RecommendedFix) -and $r.Status -ne 'OK') {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine(("Recommended Fix: {0}" -f $r.RecommendedFix))
        }
    }

    $text = $sb.ToString()
    try {
        Set-Content -Path $dlg.FileName -Value $text -Encoding UTF8
        $statusLabel.Text = "Exported: $($dlg.FileName)"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'EDC Toolkit Beta', 'OK', 'Error') | Out-Null
        return
    }

    try { Save-TextReport -Prefix 'GUI_Export' -Text $text | Out-Null } catch { }

    $rep = New-EdcResult -Category 'Reports / Logs' -CheckName 'Last export' -Status 'OK' -Summary ("Saved to {0}" -f $dlg.FileName) -Details $text -ResultGroup 'GUI Session'
    Merge-Results -Results @($rep)
    Refresh-CategoryView -Category 'Reports / Logs'
}

# -----------------------------
# Wire buttons + tab scan buttons
# -----------------------------

$btnFullScan.add_Click({ Start-Scan -Mode 'Full Scan' -Categories $script:ScanCategories })
$btnTrouble.add_Click({ Start-Scan -Mode 'Troubleshooting' -Categories $script:ScanCategories })
$btnFix.add_Click({ Apply-RecommendFixes })
$btnExport.add_Click({ Export-ResultsReport })

$btnOpenRpt.add_Click({
    try {
        $root = Get-CurrentReportRoot
        if (-not (Test-Path -Path $root)) { New-Item -Path $root -ItemType Directory -Force | Out-Null }
        Start-Process -FilePath $root | Out-Null
        Sync-ReportFilesToResults
        Refresh-CategoryView -Category 'Reports / Logs'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'EDC Toolkit Beta', 'OK', 'Error') | Out-Null
    }
})

$btnEventVwr.add_Click({
    try { Start-Process -FilePath 'eventvwr.msc' | Out-Null } catch { }
})

$btnDevMgr.add_Click({
    try { Start-Process -FilePath 'devmgmt.msc' | Out-Null } catch { }
})

$btnHelp.add_Click({
    $help = @()
    $help += "Buttons:"
    $help += "  - Run Full Scan: runs scans across all categories and captures selected outputs from the existing toolkit reports."
    $help += "  - Run Troubleshooting Checks: runs conservative health checks (OK/Warning/Failed/Not available)."
    $help += "  - Apply Recommended Fixes: applies safe fixes where available; disruptive actions require confirmation; admin may be required."
    $help += "  - Export Report: exports the current GUI results to a single text file."
    $help += "  - Open Reports Folder: opens the GUI session reports folder and refreshes the Reports/Logs tab."
    $help += "  - Open Event Viewer / Open Device Manager: quick-launches common technician tools."
    $help += "  - Ping Target / DNS Name: sets the Network ping and DNS test targets."
    $help += ""
    $help += "Tabs:"
    $help += "  - Each tab has a Scan button and a Result Group filter with icon labels."
    $help += "  - Reports / Logs shows GUI session entries and saved toolkit report files. Double-click a report to open it."
    $help += "  - Tools / Utilities checks command availability and quick-launch coverage."
    [System.Windows.Forms.MessageBox]::Show(($help -join "`r`n"), 'EDC Toolkit Beta - Help', 'OK', 'Information') | Out-Null
})

foreach ($cat in $script:AllCategories) {
    $view = $script:CategoryViews[$cat]
    $view.ScanButton.add_Click({
        $category = $this.Text -replace '^Scan\\s+',''
        if ($category -eq 'Reports / Logs') {
            Refresh-CategoryView -Category 'Reports / Logs'
            return
        }
        Start-Scan -Mode ("Scan {0}" -f $category) -Categories @($category)
    })
}

$form.add_Shown({
    $adminNote = if (Get-IsAdmin) { 'Admin: Yes' } else { 'Admin: No (re-run as Administrator for full results/fixes)' }
    $statusLabel.Text = "Ready. $adminNote"
    Sync-ReportFilesToResults
    Refresh-AllViews
})

$form.add_FormClosed({
    try { if ($logoBox.Image) { $logoBox.Image.Dispose() } } catch { }
})

if ($SelfTest) {
    try {
        $form.Show()
        [System.Windows.Forms.Application]::DoEvents()
        Start-Scan -Mode 'SelfTest' -Categories $script:ScanCategories
        $form.Close()
        'SelfTest OK'
        return
    }
    catch {
        Write-Error $_.Exception.Message
        return
    }
}

if ($SelfTestFast) {
    try {
        $form.Show()
        [System.Windows.Forms.Application]::DoEvents()
        Sync-ReportFilesToResults
        Refresh-AllViews
        [System.Windows.Forms.Application]::DoEvents()
        $form.Close()
        'SelfTestFast OK'
        return
    }
    catch {
        Write-Error $_.Exception.Message
        return
    }
}

[void]$form.ShowDialog()
