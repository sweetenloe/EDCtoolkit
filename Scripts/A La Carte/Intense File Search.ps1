Get-ChildItem C:\ -Recurse -ErrorAction SilentlyContinue |
Where-Object {$_.Name -like "*invoice*"}