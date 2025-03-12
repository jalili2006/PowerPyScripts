$msgBox = New-Object -ComObject Shell.Application
#$folder = $msgBox.BrowseForFolder(0, "Bitte Ordner wählen", 512)
$folder = $msgBox.BrowseForFile(0, "Bitte Ordner wählen", 512)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($msgBox) > $null
$src = $folder.Self.Path


#--------------------------------------------------------------------------------------------
# load excel file in powershell
$xl=New-Object -ComObject Excel.Application
$wb=$xl.WorkBooks.Open('C:\Users\User\OneDrive\Powershell\Final_Skript\eGov.csv')
$wb=$xl.WorkBooks.Open($src)
$ws=$wb.WorkSheets.item(1)
$xl.Visible=$true
#--------------------------------------------------------------------------------------------
#summerize the excel form and remove the extra columns

#DistinguishedName (Subject / Zertifikatsgegenstand)
$ws.Cells.Item(3,'J').EntireColumn.Delete()

#Anrede
#$ws.Cells.Item(3,'F').EntireColumn.Delete()

#Akad Grad
$ws.Cells.Item(3,'E').EntireColumn.Delete()

#Straße
$ws.Cells.Item(3,'H').EntireColumn.Delete()

#Bis (important for calculation)
#$ws.Cells.Item(3,'I').EntireColumn.Delete()

#Ort
$ws.Cells.Item(3,'J').EntireColumn.Delete()

#Land
$ws.Cells.Item(3,'K').EntireColumn.Delete()

#Seriennr
$ws.Cells.Item(3,'L').EntireColumn.Delete()

#Vorname
$ws.Cells.Item(3,'B').EntireColumn.Delete()

#ab
$ws.Cells.Item(3,'H').EntireColumn.Delete()

#Land
$ws.Cells.Item(3,'G').EntireColumn.Delete()

#POST
#$ws.Cells.Item(3,'N').EntireColumn.Delete()

#Email
#$ws.Cells.Item(3,'C').EntireColumn.Delete()

#Anrede
#$ws.Cells.Item(3,'D').EntireColumn.Delete()

#Strasse
$ws.Cells.Item(3,'F').EntireColumn.Delete()

#---------------------------------------------------------------------------------------------
#Add the Columns for header
$ws.Columns.ListObject.ListColumns.Add
$ws.Cells.Item(1,8) ='Fällig in ... Tagen'


#---------------------------------------------------------------------------------------------
#create a loop fill the new values in column

$i=1
Do {
    $i
    $i++
$bis = $ws.Cells.Item($i,6)

# Split the each date in Cells to easy work for calculate date and time
$gulday = $bis.text.Split(".")[0]
$gulmon = $bis.text.Split(".")[1]
$gulyer = $bis.text.Split(".")[2]



# after seperating each caracture in cells need to put on variable to prepare for calculating    
$num01 = $gulday
$num02 = $gulmon
$num03 = $gulyer

#----------------------------------------------------------------------------------------------
#use the current date to calculate day left 
$bistoday = (Get-Date -Format d)
#$ws.Cells.Item(2,11) = $bistoday

#seperate the day, month , year
$todayd= (Get-Date).Day
$todaym= (Get-Date).Month
$todayy= (Get-Date).Year


$VV01 = $num01 - $todayd
$vv02 = $num02 - $todaym
$vv03 = $num03 - $todayy

#summe all value (day,month,year) together
$vdif = $vv01 + ($vv02 * 30) + ($vv03 * 365)


#if result less than 0 means how many day go on

if ($vdif -le 0 ){

#change the color for specific cell
$ws.Cells.Item($i,8).Interior.colorIndex =23

}

ElseIf( $vdif -gt 0 -and $vdif -le 7) {

#day left
$ws.Cells.Item($i,8).Interior.colorIndex =3

#test Umlaute in excel
#$ws.Cells.Item(3,10) = "üÜÜÄÄÄÄäääääßßßßßß"
#email
$ws.Cells.Item($i,3).Interior.colorIndex =3

$email = $ws.Cells.Item($i,3).text
#$ws.Cells.Item(3,10) = $email

#Zertifikat
$zert = $ws.Cells.Item($i,7).text
$ws.Cells.Item($i,7).Interior.colorIndex =3

#BSN
$bsn = $ws.Cells.Item($i,5).text
$ws.Cells.Item($i,5).Interior.colorIndex =3

#ABLAUFDATUM
$abl = $ws.Cells.Item($i,6).text
$ws.Cells.Item($i,6).Interior.colorIndex =3


$From = "zsvu.pmo@gmail.com"
$To = $email
$Cc = "zsvu.pmo@gmail.com"
$username = "zsvu.pmo@gmail.com"

# Der E-Mail Betreff
$Subject = "Ablauf eGov-Zertifikat [" +$bsn+ "]" 
# Der E-Mail Text
$Body = "Sehr geehrte Damen und Herren,

das eGov-Zertifikat ["+ $zert +"] der ["+ $bsn +"]laeuft am ["+ $abl +"] ab. Ich bitte Sie daher um eine Verlaengerung.


Vielen Dank.

Mit freundlichen Grussen
ZSVU PMO"

$SMTPServer = "smtp.gmail.com"
$SMTPPort = "587"
#Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $username -Verbose 

}

#--------------------------------------------------------------------------------------------------------------------------------------------------

Elseif( $vdif -ge 8 -and $vdif -le 21 )  {$ws.Cells.Item($i,8).Interior.colorIndex =46}
Elseif( $vdif -ge 22 -and $vdif -le 66 )  {$ws.Cells.Item($i,8).Interior.colorIndex =6}
Else {$ws.Cells.Item($i,8).Interior.colorIndex =4}

$ws.Cells.Item($i,8) = $vdif




#Unterrichtsfreier Tag1------------------------------------------------------------------------------------------
#Unterrichtsfreier Tag Between 02.10.2017 Bis 04.10.2017

#$wins01 = '02.10.2017'
#$Wins02 = '04.10.2017'

# weis means Weinachtferien Start 
$winsd = '02'
$winsm = '10'
$winsy = '2017'

# wine means Ferien End
$wine1d = '04'
$wine2m = '10'
$wine3y = '2017'

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
#if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -gt $winsd) -and ($num01 -le '10')){

#$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"
#}


if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -ge '02') -and ($num01 -le '04')){

$ws.Cells.Item($i,9) = "Unterrichtsfreier Tag Between 02.10.2017 Bis 04.10.2017"
}


#Herbstferier------------------------------------------------------------------------------------------
#Herbstferier Tag Between 23.10.2017 Bis 04.11.2017

#$fer01 = '23.10.2017'
#$fer02 = '04.11.2017'

# fs means Ferien Start 
$fsd = 23
$fsm = 10
$fsy = 2017

# fe means Ferien End
$fed = 04
$fem = 11
$fey = 2017

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
if (($num03 -eq $fsy) -and ($num02 -eq $fsm) -and ($num01 -gt 20)){

$ws.Cells.Item($i,9) = "Herbstferier Tag Between 23.10.2017 Bis 04.11.2017"
}

Elseif (($num03 -eq $fsy) -and ($num02 -eq $fem) -and ($num01 -le '04')){

$ws.Cells.Item($i,9) = "Herbstferier Tag Between 23.10.2017 Bis 04.11.2017"

}

#Weihnachtsferien-------------------------------------------------------------------------------------
#Weihnachtsferien Tag Between 21.12.2017 Bis 02.01.2018

#$weis01 = '21.12.2017'
#$Weie02 = '02.01.2018'

# weis means Weinachtferien Start 
$weisd = 21
$weism = 12
$weisy = 2017

# fe means Ferien End
$weied = 02
$weiem = 01
$weiey = 2018

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
if (($num03 -eq $weisy) -and ($num02 -eq $weism) -and ($num01 -gt 21)){

$ws.Cells.Item($i,9) = "Weihnachtsferien Tag Between 21.12.2017 Bis 02.01.2018"
}

Elseif (($num03 -eq $weiey) -and ($num02 -eq $weiem) -and ($num01 -le '02')){

$ws.Cells.Item($i,9) = "Weihnachtsferien Tag Between 21.12.2017 Bis 02.01.2018"

}


#Winterferien------------------------------------------------------------------------------------------
#Winterferien Tag Between 05.02.2018 Bis 10.02.2018

#$wins01 = '05.02.2018'
#$Wins02 = '10.02.2018'

# weis means Weinachtferien Start 
$winsd = '05'
$winsm = '02'
$winsy = '2018'

# wine means Ferien End
$wine1d = 10
$wine2m = 02
$wine3y = 2018

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -gt $winsd) -and ($num01 -le '10')){

$ws.Cells.Item($i,9) = "Winterferien Tag Between 05.02.2018 Bis 10.02.2018"
}


#Osterferien------------------------------------------------------------------------------------------
#Osterferien Tag Between 26.03.2018 Bis 06.04.2018

#$wins01 = '26.03.2018'
#$Wins02 = '06.04.2018'

# weis means Weinachtferien Start 
$winsd = 26
$winsm = 03
$winsy = 2018

# wine means Ferien End
$wine1d = 06
$wine2m = 04
$wine3y = 2018

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
#if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -gt $winsd) -and ($num01 -le '10')){

#$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"
#}



if (($num03 -eq $winsy) -and ($num02 -eq '03') -and ($num01 -ge '26')){

$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"
}

Elseif (($num03 -eq $wine3y) -and ($num02 -eq '04') -and ($num01 -le '10')){

$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"

}



#Unterrichtsfreier Tag2------------------------------------------------------------------------------------------
#Unterrichtsfreier Tag Between 30.04.2018 Bis 02.05.2018

#$wins01 = '30.04.2018'
#$Wins02 = '02.05.2018'

# weis means Weinachtferien Start 
$winsd = 30
$winsm = 04
$winsy = 2018

# wine means Ferien End
$wine1d = 02
$wine2m = 05
$wine3y = 2018

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
#if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -gt $winsd) -and ($num01 -le '10')){

#$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"
#}



if (($num03 -eq $winsy) -and ($num02 -eq '04') -and ($num01 -ge '30')){

$ws.Cells.Item($i,9) = "Unterrichtsfreier Tag Between 30.04.2018 Bis 02.05.2018"
}

Elseif (($num03 -eq $wine3y) -and ($num02 -eq '05') -and ($num01 -le '02')){

$ws.Cells.Item($i,9) = "Unterrichtsfreier Tag Between 30.04.2018 Bis 02.05.2018"

}



#Unterrichtsfreier Tag3------------------------------------------------------------------------------------------
#Unterrichtsfreier Tag Between 11.05.2018 Bis 14.05.2018

#$wins01 = '11.05.2018'
#$Wins02 = '14.05.2018'

# weis means Weinachtferien Start 
$winsd = '11'
$winsm = '05'
$winsy = 2018

# wine means Ferien End
$wine1d = 14
$wine2m = 05
$wine3y = 2018

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
#if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -gt $winsd) -and ($num01 -le '10')){

#$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"
#}


if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -ge '11') -and ($num01 -le '14')){

$ws.Cells.Item($i,9) = "Unterrichtsfreier Tag Between 11.05.2018 Bis 14.05.2018"
}



#Pfingstferien------------------------------------------------------------------------------------------
#Pfingstferien Tag Between 22.05.2018 Bis 23.05.2018

#$wins01 = '22.05.2018'
#$Wins02 = '23.05.2018'

# weis means Weinachtferien Start 
$winsd = '22'
$winsm = '05'
$winsy = 2018

# wine means Ferien End
$wine1d = 14
$wine2m = 05
$wine3y = 2018

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
#if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -gt $winsd) -and ($num01 -le '10')){

#$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"
#}


if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -ge '22') -and ($num01 -le '23')){

$ws.Cells.Item($i,9) = "Pfingstferien Tag Between 22.05.2018 Bis 23.05.2018"
}

#------------------------------------------------------------------------------------------



#Sommerferien------------------------------------------------------------------------------------------
#Sommerferien Tag Between 04.07.2018 Bis 20.08.2018

#$wins01 = '04.07.2018'
#$Wins02 = '20.08.2018'

# weis means Weinachtferien Start 
$winsd = '04'
$winsm = '07'
$winsy = '2018'

# wine means Ferien End
$wine1d = '20'
$wine2m = '08'
$wine3y = '2018'

#if($num03 -eq 2017){ 
#if($num02 -eq 10 -and $num02 -eq 11){}
#if($num01 -eq 12 -or $num01 -eq 02){}
#$ws.Cells.Item($i,10) = "test"
#}


# $num2 means Gultig Bis month
#if (($num03 -eq $winsy) -and ($num02 -eq $winsm) -and ($num01 -gt $winsd) -and ($num01 -le '10')){

#$ws.Cells.Item($i,9) = "Osterferien Tag Between 26.03.2018 Bis 06.04.2018"
#}



if (($num03 -eq $winsy) -and ($num02 -eq '07') -and ($num01 -ge '04')){

$ws.Cells.Item($i,9) = "Sommerferien Tag Between 04.07.2018 Bis 20.08.2018"
}

Elseif (($num03 -eq $wine3y) -and ($num02 -eq '08') -and ($num01 -le '20')){

$ws.Cells.Item($i,9) = "Sommerferien Tag Between 04.07.2018 Bis 20.08.2018"

}


#------------------------------------------------------------------------------------------------------



#Sort

$xlAscending = 1
$r = $ws.UsedRange
#const xlyes 
$r2 = $ws.Range('H2') # Sorts on Column H how many day left 
$a = $r.sort($r2,1,$null,$null,1,$null,1,1)


#$ws.Cells.Item(4,10) = $r
#$ws.Cells.Item(5,10) = $r2
#$ws.Cells.Item(6,10) = $a
    
    $C_Row = $ws.UsedRange

#$a1 = $objRange.SpecialCells(11).row

$a11 = $C_Row.SpecialCells(11).row

#Write-Host "this is a test", $a11

$last_row = $a11 -1
#Write-Host "this is a test", $last_row



   }






While ($i -le $last_row)

#$Sort001 = $ws.Range("H2:$last")

#$ws.sort.setRange($sort001)

#foreach ( $i in $ws.Cells.Item($i,6) ,$i++ ) {
#$today3 = $todayd + $todaym + $todayy
#$ws.Cells.Item(2,9) = $today3
#$test001 = $ws.Cells.Item(2,9).text 
#$test001 =$test001 - 10
#$ws.Cells.Item(2,10) = $test001
#$result = $today3 - $bis
#$ws.Cells.Item(2,11) = $result