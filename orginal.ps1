#Parameter aus dem Dashboard
#Name der VM
$strVMName = Get-VM;
#Snapshotalter in Tagen
$strDays = $args[1];
 
#Suche nach Snapshots ueber gesetzte Parameter
$arroSnapshot = Get-VMSnapshot -VMName VM-Medistar;
 
#Abfrage aktuelles Datum
$dtToday = Get-Date;

#Kontrollvariable fuer Alarmierung am Skriptende; wenn $TRUE und Alarmierung aktiviert, Skriptende mit Exit 1001
$bSnapshotCheck = $FALSE;
 
#einzelne Snapshots aus Objekt-Array $arroSnapshot mit resp. Daten an das Dashboard ausgeben
foreach ($arroSnapshot in $arroSnapshot)
{
    $dtDifference = $dtToday - $arroSnapshot[0].CreationTime;
    if ($dtDifference.Days -gt $strDays)
    {
        Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
        $strOutput = "Es wurde der Snapshot " + $arroSnapshot.Name
        + " gefunden. Der Snapshot ist " + $dtDifference.Days
        + " Tage alt.";
        Write-Host $strOutput;
        Write-Host "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
        $bSnasphotCheck = $TRUE;
    }
}
 
#Pruefen ob alarmiert werden soll
if ($bSnapshotCheck -eq $TRUE)
{
    Exit 1001;
}
else
{
    #Benachrichtigung fuer Benutzer erstellen und Ausgabe an das Dashboard wenn kein Image entdeckt wurde
    Write-Host "Es wurde kein Snapshot gefunden, der aelter als " $args[1] " Tage ist.";
    Exit 0;
}

