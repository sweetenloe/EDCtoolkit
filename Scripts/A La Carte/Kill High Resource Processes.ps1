Get-Process | Sort CPU -Descending | Select -First 5 | ForEach-Object {
    Stop-Process -Id $_.Id -Force
}