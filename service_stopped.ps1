# ...existing code...
# Gets stopped services
$services = Get-Service | Where-Object { $_.Status -eq 'Stopped' } | Select-Object Name, DisplayName, Status

if (-not $services -or $services.Count -eq 0) {
    Write-Output "No stopped services found."
}

# show as grid view if available, otherwise print table
if (Get-Command -Name Out-GridView -ErrorAction SilentlyContinue) {
    $services | Out-GridView -Title 'Stopped Services'
} else {
    $services | Format-Table -AutoSize
}
# ...existing code...