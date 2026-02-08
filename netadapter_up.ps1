
get-NetAdapter | Where-Object { $_.status -eq 'Up' } 

Get-NetAdapter | Where-Object { $_.status -eq 'Up' } | Select-Object Name, Status, MacAddress, LinkSpeed

Get-NetAdapter | Select-Object ifIndex, Name

# Show only Ethernet adapters that are up, with key properties
Get-NetAdapter |
Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -like '*Ethernet*' } |
Select-Object ifIndex, Name, Status, MacAddress, LinkSpeed

# ...existing code...
Get-NetAdapter |
Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -like '*Ethernet*' } |
ForEach-Object {
    $ip = (Get-NetIPAddress -InterfaceIndex $_.ifIndex -AddressFamily IPv4 | Select-Object -ExpandProperty IPAddress)
    [PSCustomObject]@{
        ifIndex    = $_.ifIndex
        Name       = $_.Name
        Status     = $_.Status
        MacAddress = $_.MacAddress
        LinkSpeed  = $_.LinkSpeed
        IPv4       = $ip -join ', '
    }
}
# ...existing code...
# Show only Ethernet adapters that are up, with key properties and IP summary
$adapters = Get-NetAdapter |
Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -like '*Ethernet*' }

# This command retrieves a specific network adapter by its interface index (ifIndex).
# You can find the ifIndex by running the command from line 6: Get-NetAdapter | Select-Object ifIndex,Name
# Replace '15' with the actual ifIndex you want to query.
Get-NetAdapter -ifIndex 15 

<#


foreach ($adapter in $adapters) {
    $info = [PSCustomObject]@{
        ifIndex    = $adapter.ifIndex
        Name       = $adapter.Name
        Status     = $adapter.Status
        MacAddress = $adapter.MacAddress
        LinkSpeed  = $adapter.LinkSpeed
        IPs        = (Get-NetIPAddress -InterfaceIndex $adapter.ifIndex | Select-Object -ExpandProperty IPAddress)
    }
    $info
}

#>
$adapters = Get-NetAdapter

foreach ($adapter in $adapters) {
    if ($adapter.Status -eq 'Up' -and $adapter.InterfaceDescription -like '*Ethernet*') {
        # Do something with the adapter, e.g. select properties
        $adapter | Select-Object ifIndex, Name, Status, MacAddress, LinkSpeed
    }
}


# ...existing code...
Get-NetAdapter |
Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -like '*Ethernet*' } |
ForEach-Object {
    $ip = (Get-NetIPAddress -InterfaceIndex $_.ifIndex -AddressFamily IPv4 | Select-Object -ExpandProperty IPAddress)
    [PSCustomObject]@{
        ifIndex    = $_.ifIndex
        Name       = $_.Name
        Status     = $_.Status
        MacAddress = $_.MacAddress
        LinkSpeed  = $_.LinkSpeed
        IPv4       = $ip -join ', '
    }
}
# ...existing code...