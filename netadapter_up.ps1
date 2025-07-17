
get-NetAdapter | Where-Object {$_.status -eq 'Up'} 

Get-NetAdapter| Where-Object {$_.status -eq 'Up'} | Select-Object Name, Status, MacAddress, LinkSpeed

Get-NetAdapter | Select-Object ifIndex,Name

# Show only Ethernet adapters that are up, with key properties
Get-NetAdapter |
    Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -like '*Ethernet*' } |
    Select-Object ifIndex, Name, Status, MacAddress, LinkSpeed


# Show only Ethernet adapters that are up, with key properties and IP summary
$adapters = Get-NetAdapter |
    Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -like '*Ethernet*' }

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