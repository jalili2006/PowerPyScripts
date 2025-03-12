$dat = Get-VMSnapshot -VMName VM-Medistar;
Get-VMSnapshot -VMName VM-Medistar | sort-object -property @{Expression={$_.LastWriteTime - $_.CreationTime}; Ascending=$true}


foreach ($dat in $dat){

$tim = $dat.CreationTime
  
    $d2 = $tim.Day
    $m2 = $tim.Month
    $y2 = $tim.Year
    $final2 = $d2 + $m2 +$y2

   
  #Convert current date to the integer 
      
    Write-Host "-----------------------------";
    $d = Get-Date 
    $d1=$d.Day
    $m1=$d.Month
    $y1=$d.Year
    $final1= $d1 + $m1 + $y1


# find differentiate between two dates


$diff1= $final1 - $final2


#Write-Host $final1;
        
#Write-Host $final2;

#Write-Host $diff1;

if ($diff1 -ge 3) {

Write-Host "Creation of VM Snapshot is later than 3 days " ":" $diff1 "day";
Write-Host "Snapshot Creation Time" :$tim;
Write-Host "Snapshot Name" : $Nsnap = $dat.Name;
#Exit 1001;


}

else {
 
Write-Host "Creation of VM Snapshot is less than 3 days" ":" $diff1 "day"; 
Write-Host "Snapshot Creation Time" :$tim;
Write-Host "Snapshot Name" : $Nsnap = $dat.Name;
#Exit 0;

}

Write-Host $diff1 

#$dat.CreationTime |Sort-Object 

#$tim |Sort -Descending


}

