#Enter-PSSession -ComputerName 192.168.120.20 -Credential Administrator
#Dism /online /enable-feature /FeatureName:Microsoft-Hyper-V
#Rename-Computer -ComputerName Medistar -NewName nano -LocalCredential administrator -Restart
#Rename-Computer -ComputerName Medistar -NewName nano -LocalCredential ex\administrator -Restart
#____________________________________________________________________________________________
##(1)Create blob file in any localmachine
#djoin /PROVISION /DOMAIN ex.local /MACHINE NANO /SAVEFILE C:\nano\blob001.txt


##(2)After create local blob from another machine and load with below command:
#djoin /requestodj /loadfile C:\nano\blob.txt /windowspath C:\Windows /localos

#____________________________________________________________________________________________
#Restart-Computer -ComputerName nano

#____________________________________________________________________________________________
#Get-DnsClientServerAddress -InterfaceAlias Ethernet -InterfaceIndex 2 -AddressFamily IPv4;
#Set-DnsClientServerAddress -InterfaceAlias Ethernet -ServerAddresses ("192.168.120.1")

#____________________________________________________________________________________________
#Install Hyper-v over Nano server

 









