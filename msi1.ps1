#input remote computer name
$ComputerName = Read-Host "Name of the computer being remote accessed"

#start WinRm service if stopped
Get-Service WinRM -ComputerName $ComputerName | Where {$_.status –eq 'Stopped'} |  Start-Service

#Create a folder on remote computer
New-Item -ItemType directory -Path "\\$ComputerName\C$\temp"

#Copy the msi to the remote coputer
xcopy C:\temp\winzip260-64.msi \\$ComputerName\C$\temp

#Pause 3 sec.
Start-Sleep -s 3

#Message that he managed to copy the file to the remote computer in yellow
write-host "Winzip.msi has been copied to the remote computer" -ForegroundColor Yellow

#Pause 3 sec.
Start-Sleep -s 3

#Install on the remote computer
Invoke-Command -ComputerName $ComputerName -ScriptBlock {&cmd.exe /c MSIEXEC /I "c:\temp\winzip260-64.msi" /qn}

#Message that he managed to install the file to the remote computer in green
write-host "Winzip.msi software has been installed on the remote computer" -ForegroundColor Green
