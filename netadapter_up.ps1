
get-NetAdapter | Where-Object {$_.status -eq 'Up'} 

Get-NetAdapter| Where-Object {$_.status -eq 'Up'} | Select-Object Name, Status, MacAddress, LinkSpeed

Get-NetAdapter | Select-Object ifIndex,Name

