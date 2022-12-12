$menu=@"
1 Get Windows Infos for installation (Hostname,IP,OS)
2 Get NTCS Infos for installation
3 Set Power Option befor installation (Höchstleistung)
4 Set NTCS Option befor installation (PendingFileRenameOperations)
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
        $hostname = $env:COMPUTERNAME #Get the hostname - 1st try
        if ($hostname -eq $null -or $hostname.Count -eq 0){$hostname = hostname} #Get hostname - 2nd try
        if ($hostname -eq $null -or $hostname.Count -eq 0){$hostname = gc env:computername} #Get hostname - 3nd try

        ##Get the OS name and version
        $os = $null
        $os = Get-ComputerInfo -Property "OSName", "OSVersion" #Get the OS name and version - 1st try
        if ($os -eq $null -or $os.Count -eq 0){
            $os = (Get-WmiObject win32_operatingsystem).caption
            $os += (Get-WmiObject win32_operatingsystem).Version

        } #Get the OS name - 2nd try

        #Get the IP address
        $ipAddress = Get-NetIPAddress | Where-Object {$_.AddressFamily -eq "IPv4"} | Select-Object -First 1 #Get the IP address 1st try
        if ($ipAddress -eq $null -or $ipAddress.Count -eq 0){$ipAddress = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE).IPAddress[0]} # Get the IP address 2nd try  

        #Get the antivirus product
        $antivirusProduct = $null
        $Virenscanners = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct -ErrorAction SilentlyContinue
            foreach($Virenscanner in $Virenscanners){
                $antivirusProduct += $Virenscanner.displayName
            } #Get the antivirus product 1st try
        if ($antivirusProduct -eq $null -or $antivirusProduct.Count -eq 0){
            $antivirusProdu = Get-MpComputerStatus -ErrorAction SilentlyContinue
            $antivirusProduct = $antivirusProdu.AntivirusProduct 

        } #Get the antivirus product 2st try
        if ($antivirusProduct -eq $null -or $antivirusProduct.Count -eq 0){
            $wmiQuery_avp = "SELECT * FROM AntiVirusProduct" 
            $AntivirusProduct_ = Get-WmiObject -Namespace "root\SecurityCenter2" -Query $wmiQuery_avp  @psboundparameters -ErrorVariable myError -ErrorAction 'SilentlyContinue'             
            $antivirusProduct = $AntivirusProduct_.displayName
        } #Get antivirus product 3nd try
        if ($antivirusProduct -eq $null -or $antivirusProduct.Count -eq 0){$antivirusProduct = Get-MpComputerStatus -ErrorAction SilentlyContinue} #Get antivirus product 4nd try

        # Check if SQL Server is installed
        $sqlServerInstalled = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Microsoft SQL Server" -ErrorAction SilentlyContinue
        if ($sqlServerInstalled) {
            $sqlServerInstalled_String = "SQL Server is installed on this machine"

            #Get installed SQL Versions
            $SQLVersion_regs = (get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances
            foreach ($SQLVersion_reg in $SQLVersion_regs){
               $SQLVersion_ = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$SQLVersion_reg
               $SQLVersion_main = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$SQLVersion_\Setup").Edition
               $SQLVersion_sub = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$SQLVersion_\Setup").Version
            }
        } else {
            $sqlServerInstalled_String = "SQL Server is not installed on this machine"
        }

        ### Output
        # Display the hostname, OS, IP address, and antivirus product
        Write-Host "Hostname: $hostname" -ForegroundColor Yellow
        Write-Host "OS: $($os.OSName) $($os.OSVersion)" -ForegroundColor Yellow
        Write-Host "IP Address: $($ipAddress.IPAddress)" -ForegroundColor Yellow
        if ($antivirusProduct -eq $null -or $antivirusProduct.Count -eq 0) {
           Write-Host "Antivirus Product: not found! Windows Server?" -ForegroundColor Red
        }else{
            Write-Host "Antivirus Product: $antivirusProduct" -ForegroundColor Yellow
            }
        Write-Host "SQL: $sqlServerInstalled_String" -ForegroundColor Yellow
        if ($sqlServerInstalled_String -like "SQL Server is installed on this machine"){Write-Host "SQL Version:"$SQLVersion_main $SQLVersion_sub -ForegroundColor Yellow}
        pause
         } 

     "2" {
        Write-Host "1. NTCS Standard Path (x86) press 1" 
        Write-Host "2. custom NTCS Path press 2"

        # Get user input
        $input = Read-Host

        # Process user input
        if ($input -eq "1") {
          # Perform actions for option 1 - NTCS Version mit Standard Pfad
          $rootPath_NTCS = "C:\Program Files (x86)\BMDSoftware\BMDNTCS.exe"
          if (Test-Path -path $rootPath_NTCS){
          $NTCSversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($rootPath_NTCS).FileVersion
          Write-Host "NTCS Version: "$NTCSversion -ForegroundColor Yellow
          }else{
          Write-Host "NTCS is not installed on $rootPath_NTCS <=> Check Paths!!" -ForegroundColor Red
          }
        } elseif ($input -eq "2") {
          # Perform actions for option 2 - NTCS Version mit Custom Pfad
          $rootPath_NTCS = ""
          $rootPath_NTCS = Read-Host "Pfad für NTCS eingeben: " 
          $rootPath_NTCS += "\BMDNTCS.exe"
          if (Test-Path -path $rootPath_NTCS){
          $NTCSversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($rootPath_NTCS).FileVersion
          Write-Host "NTCS Version: "$NTCSversion -ForegroundColor Yellow
          }else{
          Write-Host "NTCS is not installed on $rootPath_NTCS <=> Check Paths!!" -ForegroundColor Red
          }
        } else {
          # Handle invalid input
          Write-Warning "Invalid Choice. Try again."
              sleep -milliseconds 750
        }

        pause
          }
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
        }
        }
        pause
        }
    "4" {
        #Regwert ändern PendingFileRenameOperations

        if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name _PendingFileRenameOperations -ErrorAction SilentlyContinue) {
        Write-Host 'Info: found renamed _PendingFileRenameOperations - not necessary? check manuel' -ForegroundColor Red
            #Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\" -name "_PendingFileRenameOperations" -ErrorAction SilentlyContinue
        Return
        }

        if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) {
           $confirmation = Read-Host "rename PendingFile - Do you want to continue? [1 to continue] "
            if ($confirmation -eq '1') {
                rename-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name "PendingFileRenameOperations" -NewName "_PendingFileRenameOperations" -ErrorAction SilentlyContinue
                if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) {Write-Host "keine Rechte? Skript als Admin starten" -ForegroundColor Red}
                if (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name _PendingFileRenameOperations -ErrorAction SilentlyContinue) {Write-Host "RegFile: PendingFileRenameOperations wurde angepasst" -ForegroundColor Yellow}
            }
        }else{
            Write-Host 'PendingFileRenameOperations NOT exist - not necessary? check manuel' -ForegroundColor Red
        }
        pause
         }
     "Q" {Write-Host "Goodbye by ma1r1" -ForegroundColor Cyan
         Return
         }
     Default {Write-Warning "Invalid Choice. Try again."
              sleep -milliseconds 750}
    } #switch
} While ($True) #make the action
##End-Script