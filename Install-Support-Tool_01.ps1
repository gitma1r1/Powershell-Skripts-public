$menu=@"
1 Get Windows Infos (Hostname,IP,OS)
2 Get BMDNTCS Infos (for installation)
3 Set Höchstleistung
4 Set PendingFileRenameOperations
Q Quit
 
Select a task by number or Q to quit
"@ #def. var. menu

Function Invoke-Menu {
    [cmdletbinding()]
    
    Param(
        [Parameter(Position=0,Mandatory=$True,HelpMessage="Enter your menu text")]
        [ValidateNotNullOrEmpty()]
        [string]$Menu,
        [Parameter(Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$Title = "My Menu",
        [Alias("cls")]
        [switch]$ClearScreen
    )

    if ($ClearScreen) {Clear-Host}

$menuPrompt = $title #build the menu prompt
$menuprompt+="`n" #add a return
$menuprompt+="-"*$title.Length #add an underline
$menuprompt+="`n" #add another return
$menuPrompt+=$menu #add the menu

Read-Host -Prompt $menuprompt

} #end function for the menu

Do {
    Switch (Invoke-Menu -menu $menu -title "Install-Support-Tool by mai156" -clear) { #use a Switch construct to take action depending on what menu choice is selected.
    
     "1" {

        ##Get the hostname
        function get-hostname{
            $global:hostname = $null
            #1st try - get-hostname
            $global:hostname = $env:COMPUTERNAME #Get the hostname - 1st try
            #2st try - get-hostname
            if ($global:hostname -eq $null -or $global:hostname.Count -eq 0){$global:hostname = hostname} #Get hostname - 2nd try
            #3st try - get-hostname
            if ($global:hostname -eq $null -or $global:hostname.Count -eq 0){$global:hostname = gc env:computername} #Get hostname - 3nd try
        }get-hostname

        ##Get the OS name and version
        function get-OS{
            $global:os = $null
            #1st try - get-OS
            $global:os = Get-ComputerInfo -Property "OSName", "OSVersion" #Get the OS name and version - 1st try
            #2st try - get-OS
            if ($global:os -eq $null -or $global:os.Count -eq 0){
                $global:os = (Get-WmiObject win32_operatingsystem).caption
                $global:os += (Get-WmiObject win32_operatingsystem).Version

            } #Get the OS name - 2nd try
        }get-OS

        #Get the IP address
        function get-IpAdress{
            #1st try - get-OS
            $global:ipAddress = Get-NetIPAddress | Where-Object {$_.AddressFamily -eq "IPv4"} | Select-Object -First 1 #Get the IP address 1st try
            #2st try - get-OS
            if ($global:ipAddress -eq $null -or $global:ipAddress.Count -eq 0){$global:ipAddress = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE).IPAddress[0]} # Get the IP address 2nd try  
        }get-IpAdress

        #Get the public IP address
        function get-publicIP{
            #1st try - get-publicIP
            $global:publicIP = (Invoke-WebRequest -uri "http://ifconfig.me/ip").Content #1st try - get-publicIP
        }get-publicIP

        #Get the antivirus product
        function get-antivirusProduct{
            $global:antivirusProduct = $null
            #1st try - get-antivirusProduct
            $Virenscanners = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct -ErrorAction SilentlyContinue
                foreach($Virenscanner in $Virenscanners){
                    $global:antivirusProduct += $Virenscanner.displayName
                } #Get the antivirus product 1st try
            #2st try - get-antivirusProduct
            if ($global:antivirusProduct -eq $null -or $global:antivirusProduct.Count -eq 0){
                $antivirusProdu = Get-MpComputerStatus -ErrorAction SilentlyContinue
                $global:antivirusProduct = $antivirusProdu.AntivirusProduct 
            } #Get the antivirus product 2st try
            #3st try - get-antivirusProduct
            if ($global:antivirusProduct -eq $null -or $global:antivirusProduct.Count -eq 0){
                $wmiQuery_avp = "SELECT * FROM AntiVirusProduct" 
                $AntivirusProduct_ = Get-WmiObject -Namespace "root\SecurityCenter2" -Query $wmiQuery_avp  @psboundparameters -ErrorVariable myError -ErrorAction 'SilentlyContinue'             
                $global:antivirusProduct = $AntivirusProduct_.displayName
            } #Get antivirus product 3nd try
            #4st try - get-antivirusProduct
            if ($global:antivirusProduct -eq $null -or $global:antivirusProduct.Count -eq 0){$global:antivirusProduct = Get-MpComputerStatus -ErrorAction SilentlyContinue} #Get antivirus product 4nd try
        }get-antivirusProduct

        # Check if SQL Server is installed
        function get-sqlInstalled{
            #1st try - get-sqlInstalled
            $global:sqlServerInstalled = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Microsoft SQL Server" -ErrorAction SilentlyContinue
            if ($global:sqlServerInstalled) {
                $sqlServerInstalled_String = "SQL Server is installed on this machine"
                #Get installed SQL Versions
                $global:SQLVersion_regs = (get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances
                    foreach ($SQLVersion_reg in $SQLVersion_regs){
                        $SQLVersion_ = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$SQLVersion_reg
                        $SQLVersion_main = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$SQLVersion_\Setup").Edition
                        $SQLVersion_sub = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$SQLVersion_\Setup").Version
                    }
                }else{
                $sqlServerInstalled_String = "SQL Server is not installed on this machine"
                }
        }get-sqlInstalled

        function check-admin-access{
            $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal $currentUser
            $admin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

            if ($admin) {
              $global:adminAccess = "YES! You have administrative access."
            } else {
              $global:adminAccess = "NO! You do not have administrative access."
            }
        }check-admin-access


        ### Output
        # Display Basic Infos
        Write-Host "Hostname: " -BackgroundColor black -NoNewline;
        Write-Host $global:hostname -ForegroundColor green -BackgroundColor black 
        Write-Host "OS: " -BackgroundColor black -NoNewline;
        Write-Host "$($global:os.OSName) $($global:os.OSVersion)" -ForegroundColor green -BackgroundColor black
        Write-Host "IP Address: "-BackgroundColor black -NoNewline;
        Write-Host "$($global:ipAddress.IPAddress)" -ForegroundColor green -BackgroundColor black
        if ($global:antivirusProduct -eq $null -or $global:antivirusProduct.Count -eq 0) {
           Write-Host "Antivirus Product: not found! Windows Server?" -ForegroundColor Red
        }else{
            Write-Host "Antivirus Product: " -BackgroundColor black -NoNewline;
            Write-Host "$global:antivirusProduct" -ForegroundColor green -BackgroundColor black
            }
        Write-Host "SQL: " -BackgroundColor black -NoNewline;
        Write-Host "$sqlServerInstalled_String" -ForegroundColor green -BackgroundColor black
        if ($sqlServerInstalled_String -like "SQL Server is installed on this machine"){Write-Host "SQL Version:"$SQLVersion_main $SQLVersion_sub -ForegroundColor green -BackgroundColor black}
        Write-Host ""
        Write-Host "public IP: " -BackgroundColor black -NoNewline;
        Write-Host $publicIP -ForegroundColor green -BackgroundColor black        
        Write-Host "Admin Access: " -BackgroundColor black -NoNewline;
        Write-Host $global:adminAccess -ForegroundColor green -BackgroundColor black
        
        pause
         }# end menu1

     "2" {

        Write-Host "a. Check if NTCS is installed on Basic Paths..." -ForegroundColor Yellow
        $counter = 0
        $global:rootPath_NTCS = "C:\Program Files (x86)\BMDSoftware\BMDNTCS.exe"
        $global:Path_NTCS_01 = "D:\Program Files (x86)\BMDSoftware\BMDNTCS.exe"
        $global:Path_NTCS_02 = "D:\Programme\BMDSoftware\BMDNTCS.exe"
        $global:Path_NTCS_03 = "D:\BMDNTCS\BMDSoftware\BMDNTCS.exe"
        $global:Path_NTCS_04 = "D:\BMD\BMDNTCS.exe"

        $check_Paths_NTCS = Test-Path -Path $rootPath_NTCS , $Path_NTCS_01 , $Path_NTCS_02 , $Path_NTCS_03 , $Path_NTCS_04

        foreach($check_Path_NTCS in $check_Paths_NTCS){
            if($check_Path_NTCS -eq $True){
            #check NTCS exe
            Write-Host "NTCS is installed check standards"
            }else{
            $counter ++
            }
        }

        if ($counter -gt 0){
          Write-Host "b. NTCS auf " -NoNewline;  
          Write-Host $counter -ForegroundColor Red -NoNewline;
          Write-Host " Pfaden gesucht: " -NoNewline;
          Write-Host "nicht gefunden" -ForegroundColor Red -NoNewline;
          Write-Host " - manuell eingeben!"
          }

          $manuel_Path_NTCS = ""
          $manuel_Path_NTCS = Read-Host "c. Pfad für NTCS eingeben: " 
          $manuel_Path_NTCS += "\BMDNTCS.exe"
          if (Test-Path -path $manuel_Path_NTCS){
            $NTCSversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($manuel_Path_NTCS).FileVersion
            Write-Host "NTCS Version: "$NTCSversion -ForegroundColor Yellow
          }else{
            $counter ++
            Write-Host "d. NTCS is not installed on $manuel_Path_NTCS - $counter NTCS Paths checked" -ForegroundColor Red
          }

        pause
          } #end menu 2


     "3" {

        $getactivePowerPlan = PowerCfg.exe /GETACTIVESCHEME
        if ($getactivePowerPlan -like "*8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c*"){
            Write-Host Windows Energieoptionen sind schon auf Höchstleistung gesetzt -ForegroundColor Yellow
        }else{
        $confirmation = Read-Host "Energieoptionen Höchstleistung setzen - Do you want to continue? [1 to continue] " 
            if ($confirmation -eq '1') {
        powercfg -SETACTIVE 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 
        $getactivePowerPlan_new = PowerCfg.exe /GETACTIVESCHEME
        Write-Host "Windows Energieoptionen wurde auf $getactivePowerPlan_new gesetzt" -ForegroundColor Yellow
        }else {
          # Handle invalid input
          Write-Warning "Invalid Choice. Try again."
              sleep -milliseconds 750
        }
        }
        pause
        } #end menu 3

    "4" {

        #Regwert ändern PendingFileRenameOperations

        if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name _PendingFileRenameOperations -ErrorAction SilentlyContinue) {
            Write-Host 'Info: found renamed _PendingFileRenameOperations - not necessary? check manuel' -ForegroundColor Red
            #Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\" -name "_PendingFileRenameOperations" -ErrorAction SilentlyContinue
        #Return
        }

        if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) {
           $confirmation = Read-Host "rename PendingFile - Do you want to continue? [1 to continue] "
            if ($confirmation -eq '1') {
                rename-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name "PendingFileRenameOperations" -NewName "_PendingFileRenameOperations" -ErrorAction SilentlyContinue
                if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) {Write-Host "keine Rechte? Skript als Admin starten" -ForegroundColor Red}
                if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name _PendingFileRenameOperations -ErrorAction SilentlyContinue) {Write-Host "RegFile: PendingFileRenameOperations wurde angepasst" -ForegroundColor Yellow}
            }else {
          # Handle invalid input
          Write-Warning "Invalid Choice. Try again."
              sleep -milliseconds 750
        }
        }else{
            Write-Host 'PendingFileRenameOperations NOT exist - not necessary? check manuel' -ForegroundColor Red 
        }
        pause

         }#end Menu 4


     "Q" {Write-Host "Goodbye by ma1r1" -ForegroundColor Cyan
         Return
         }
     Default {Write-Warning "Invalid Choice. Try again."
              sleep -milliseconds 750}
    } #switch
} While ($True) #make the action
##End-Script