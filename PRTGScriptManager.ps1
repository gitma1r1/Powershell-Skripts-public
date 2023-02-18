#############
#Description: PRTG Script Manager
#
#Created by: Martin Mairinger
#Created on: 01.02.23
#
#Edited by: Martin Maiirnger
#Edited on: 01.02.23
#
#Changelog: Version 1
#############


##########################################################################################################
#XAML Part Begin

    Add-Type -AssemblyName PresentationFramework

    # where is the XAML file?
    $xamlFile = "C:\bin\PRTGScriptRepository\PRTGScriptManager\MainWindows.xaml"

    #create window
    $inputXML = Get-Content $xamlFile -Raw
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [XML]$XAML = $inputXML

    #Read XAML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    try {
        $window = [Windows.Markup.XamlReader]::Load( $reader )
    } catch {
        Write-Warning $_.Exception
        throw
    }

    # Create variables based on form control names.
    # Variable will be named as 'var_<control name>'

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        #"trying item $($_.Name)"
        try {
            Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
        } catch {
            throw
        }
    }

#XAML Part End
##########################################################################################################
#Your Code start - #Get-Variable var_*

#global vars
$global:stablePath = "C:\bin\PRTGScriptRepository\Stable"
$global:devPath = "C:\bin\PRTGScriptRepository\Dev"
$global:oldPath = "C:\bin\PRTGScriptRepository\Old"
$global:TemplatePs1 = "C:\bin\PRTGScriptRepository\Templates\BMDPRTGSensorTemplate.ps1"


#disable the copy buttons befor select button runs
$var_btn_copy_to_stable.isEnabled = $false
$var_btn_copy_to_dev.isEnabled = $false
#disable the var_bt_run_in_ise
$var_bt_run_in_ise.isEnabled = $false
#disable the clear dev buttons befor
$var_bt_clear_dev_folder.IsEnabled = $false
#disable the new script buttons befor
$var_btn_new_sensor.IsEnabled = $false


#Functions

#Function CopyScriptDev
Function CopyScriptDev {
    param (
        [string]$srcFile,
        [string]$destFile,
        [string]$copyToDevButtonText
        #[string]$new_Sensor_Name,
        ##[string]$new_Sensor_Desc,
        #[string]$new_Sensor_Creator,
        #[string]$new_Sensor_ChangeLog
    )

    #$new_Sensor_Name = $var_btn_new_sensor.text.ToString()
    ##$new_Sensor_Desc = $var_tb_new_sensor_desc.text.ToString()
    #$new_Sensor_Creator = $var_tb_new_sensor_creator.text.ToString()
    #$new_Sensor_ChangeLog = $var_tb_new_sensor_name_changlog.text.ToString()


    if (Test-Path $destFile) {

        $continue = [System.Windows.MessageBox]::Show("The file $destFile already exists. Do you want to replace it?", "PRTG Script Manager", 'YesNo');

        if ($continue -eq 'Yes'){
            Copy-Item $srcFile $destFile
            $contents = Get-Content $destFile
            $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
            $newLine = "# Copied to Dev on $date by $copyToDevButtonText"
            $contents = Get-Content -Path $destFile
            $contentsWithLineBreaks = $contents -join [Environment]::NewLine
            Set-Content -Path $destFile -Value ($newLine + [Environment]::NewLine + $contentsWithLineBreaks)
            [System.Windows.MessageBox]::Show("'$srcFile' copied  to '$destFile' (replaced)", 'PRTG Script Manager Message')
            Write-Host "File '$srcFile' copied to '$destFile' (replaced)" -ForegroundColor Yellow
        }
        if ($continue -eq 'No'){
            [System.Windows.MessageBox]::Show('nothing happend!', 'PRTG Script Manager Message')
            Write-Host "nothing happend!"
            return
        }
    }else{
        Copy-Item $srcFile $destFile
        $contents = Get-Content $destFile
        $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
        $newLine = "# Copied to Dev on $date by $copyToDevButtonText"
        $contents = Get-Content -Path $destFile
        $contentsWithLineBreaks = $contents -join [Environment]::NewLine
        Set-Content -Path $destFile -Value ($newLine + [Environment]::NewLine + $contentsWithLineBreaks)
        Set-Content $destFile ($newLine + [Environment]::NewLine + $contentsWithLineBreaks)
        [System.Windows.MessageBox]::Show("'$srcFile' copied  to '$destFile'", 'PRTG Script Manager Message')
        Write-Host "File '$srcFile' copied  to '$destFile'" -ForegroundColor Yellow
    }
}


#Button: "get script from stable"
$var_btn_get_script_from_stable.Add_Click( {
$window.TopMost = $false
    $global:StableScriptFiles = Get-ChildItem $stablePath
    $var_lb_selected_scripts.Items.Clear()
    foreach($file in $StableScriptFiles){
        $var_lb_selected_scripts.Items.Add($file.Name)
    }
    $var_btn_copy_to_stable.isEnabled = $false 
    $var_btn_copy_to_dev.isEnabled = $true
    $var_bt_run_in_ise.isEnabled = $true
    $var_lb_selected_scripts.SelectedIndex = 0
 
   })
   
#Button: "get script from dev"
$var_bt_get_script_from_dev.Add_Click({
$window.TopMost = $false
   $global:DevScriptFiles = Get-ChildItem $devPath
    $var_lb_selected_scripts.Items.Clear()
    foreach($file in $DevScriptFiles){
        $var_lb_selected_scripts.Items.Add($file.Name)
    }
    $var_btn_copy_to_stable.isEnabled = $true
    $var_btn_copy_to_dev.isEnabled = $false
    $var_bt_run_in_ise.isEnabled = $true
    $var_lb_selected_scripts.SelectedIndex = 0
})


# Event handler for when the selection in the list box changes
$var_lb_selected_scripts.Add_SelectionChanged({
$window.TopMost = $false
    $selectedScript = $var_lb_selected_scripts.SelectedItem
    if ($selectedScript){
        if ($selectedScript -like "dev_*") {
            $srcFile = Join-Path $devPath $selectedScript
        }else {
            $srcFile = Join-Path $stablePath $selectedScript
        }
        $content = Get-Content $srcFile -TotalCount 500
        Write-Host "selected: "$srcFile
        $var_tb_script_info.Text = $content  -join [Environment]::NewLine
        # add the selected script filename to the textbox
        $var_tb_selected_script.Text = $srcFile
    }else {
        Write-Host "new select run started"
    }
})

#Button: "copy to dev"
$var_btn_copy_to_dev.Add_Click({
$window.TopMost = $false
    $selectedScript = $var_lb_selected_scripts.SelectedItem
    if ($selectedScript.Count -eq 0){
        [System.Windows.MessageBox]::Show("No script selected" ,'PRTG Script Manager Message')
        return
    }
 
        $srcFile = Join-Path $stablePath $selectedScript
        $destFile = Join-Path $devPath "dev_$selectedScript"

    CopyScriptDev -srcFile $srcFile -destFile $destFile

})

#Button: "copy to stable"
$var_btn_copy_to_stable.Add_Click({
$window.TopMost = $false
    $var_tb_script_info.Text = "Button Copy to Stable is clicked"
    Write-Host "Button Copy to Stable is clicked"
    

    $selectedScripts = $var_lb_selected_scripts.SelectedItems
    if ($selectedScripts.Count -eq 0){
        [System.Windows.MessageBox]::Show("No script selected" ,'PRTG Script Manager Message')
        return
    }
})



#Button: "bt_run_in_ise"
$var_bt_run_in_ise.Add_Click({
$window.TopMost = $false
    $selectedScripts = $var_lb_selected_scripts.SelectedItems
    if ($selectedScripts.Count -eq 0){
        [System.Windows.MessageBox]::Show("No script selected" ,'PRTG Script Manager Message')
        return
    }

    $selectedScript = $var_lb_selected_scripts.SelectedItem
    if ($selectedScript){
        if ($selectedScript -like "dev_*") {
            $srcFile = Join-Path $devPath $selectedScript
        }else {
            $srcFile = Join-Path $stablePath $selectedScript
        }
       # $content = Get-Content $srcFile -TotalCount 12
        Write-Host "selected: "$srcFile
        PowerShell_Ise.exe -file $srcFile
        $var_tb_script_info.Text = "ISE with $srcFile opened"
        Write-Host "ISE with $srcFile opened" -ForegroundColor Yellow
    }else {
        Write-Host "new select run started"
    }
})



#Button: "bt_add_desc_changes"
$var_bt_add_desc_changes.Add_Click({
$window.TopMost = $false

    
    $newScriptDesc = $var_tb_new_sensor_desc.text.ToString()
    $newScriptCreator = $var_tb_new_sensor_creator.text.ToString()
    $newScriptChanglog = $var_tb_new_sensor_name_changlog.text.ToString()

    $StatusChechkbox_newSensor_desc = $var_cb_new_sensor_desc.IsChecked
    $StatusChechkbox_newSensor_creator = $var_cb_new_sensor_creator.IsChecked
    $StatusChechkbox_newSensor_changlog = $var_cb_new_sensor_changlog.IsChecked


    if ($StatusChechkbox_newSensor_desc -eq "true"){
    Write-Host "Desc:"$newScriptDesc
    $selectedScript = $var_lb_selected_scripts.SelectedItem
        if ($selectedScript){
        if ($selectedScript -like "dev_*") {
            $ScriptPath = Join-Path $devPath $selectedScript
        }else {
                $ScriptPath = Join-Path $stablePath $selectedScript
        }

        $contents = Get-Content $ScriptPath
        $index = [array]::IndexOf($contents, ($contents | Select-String -Pattern "#Description:").Line)

        
        if ($index -eq -1) {
            return
            Write-Host "error"
        } else {
            $contents[$index] = "#Description: $newScriptDesc"
            Set-Content -Path $ScriptPath -Value $contents
            $contentsTime = Get-Content $ScriptPath
            $index_edit_time = [array]::IndexOf($contentsTime, ($contentsTime | Select-String -Pattern "#Edited on").Line)
            $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
            $contentsTime[$index_edit_time] = "#Edited on: $date"
            Set-Content -Path $ScriptPath -Value $contentsTime
        }
    }
}



    if ($StatusChechkbox_newSensor_creator -eq "true"){
        Write-Host "Creator:"$newScriptCreator
        $selectedScript = $var_lb_selected_scripts.SelectedItem
          if ($selectedScript){
            if ($selectedScript -like "dev_*") {
                $ScriptPath = Join-Path $devPath $selectedScript
            }else {
                    $ScriptPath = Join-Path $stablePath $selectedScript
            }

            $contents = Get-Content $ScriptPath
            $index = [array]::IndexOf($contents, ($contents | Select-String -Pattern "#Edited by:").Line)
            if ($index -eq -1) {
                return
                Write-Host "error"
            } else {
                $contents[$index] = "#Edited by: $newScriptCreator"
                Set-Content -Path $ScriptPath -Value $contents
            $contentsTime = Get-Content $ScriptPath
            $index_edit_time = [array]::IndexOf($contentsTime, ($contentsTime | Select-String -Pattern "#Edited on").Line)
            $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
            $contentsTime[$index_edit_time] = "#Edited on: $date"
            Set-Content -Path $ScriptPath -Value $contentsTime
            }
        }
    }

    if ($StatusChechkbox_newSensor_changlog -eq "true"){
        Write-Host "ChangeLog:"$newScriptChanglog
        $selectedScript = $var_lb_selected_scripts.SelectedItem
          if ($selectedScript){
            if ($selectedScript -like "dev_*") {
                $ScriptPath = Join-Path $devPath $selectedScript
            }else {
                    $ScriptPath = Join-Path $stablePath $selectedScript
            }

            $contents = Get-Content $ScriptPath
            $index = [array]::IndexOf($contents, ($contents | Select-String -Pattern "#Changelog:").Line)
            if ($index -eq -1) {
                return
                Write-Host "error"
            } else {
                $contents[$index] = "#Changelog: $newScriptChanglog"
                Set-Content -Path $ScriptPath -Value $contents
            $contentsTime = Get-Content $ScriptPath
            $index_edit_time = [array]::IndexOf($contentsTime, ($contentsTime | Select-String -Pattern "#Edited on").Line)
            $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
            $contentsTime[$index_edit_time] = "#Edited on: $date"
            Set-Content -Path $ScriptPath -Value $contentsTime
            }
        }
    }
    
})









#Button: "btn_sync_baramundi"
$var_btn_sync_baramundi.Add_Click({
$window.TopMost = $false
    $var_tb_script_info.Text = "Button btn_sync_baramundi is clicked"
    Write-Host "Button btn_sync_baramundi is clicked"
    
})

#Button: "btn_new_sensor"
$var_btn_new_sensor.Add_Click({
$window.TopMost = $false
    $newScriptName = $var_tb_new_sensor_name.text.ToString() #neuer Script Name
    $newScriptDevPath = $devPath + "\dev_" + $newScriptName #Pfad



    CopyScriptDev -srcFile $TemplatePs1 -destFile $newScriptDevPath


    $newScriptDesc = $var_tb_new_sensor_desc.text.ToString()
    $newScriptCreator = $var_tb_new_sensor_creator.text.ToString()
    $newScriptChanglog = $var_tb_new_sensor_name_changlog.text.ToString()

    $StatusChechkbox_newSensor_desc = $var_cb_new_sensor_desc.IsChecked
    $StatusChechkbox_newSensor_creator = $var_cb_new_sensor_creator.IsChecked
    $StatusChechkbox_newSensor_changlog = $var_cb_new_sensor_changlog.IsChecked


    if ($StatusChechkbox_newSensor_desc -eq "true"){
    Write-Host "Desc:"$newScriptDesc

        $contents = Get-Content $newScriptDevPath
        $index = [array]::IndexOf($contents, ($contents | Select-String -Pattern "#Description:").Line)

        
        if ($index -eq -1) {
            return
            Write-Host "error"
        } else {
            $contents[$index] = "#Description: $newScriptDesc"
            Set-Content -Path $newScriptDevPath -Value $contents
            $contentsTime = Get-Content $newScriptDevPath
            $index_edit_time = [array]::IndexOf($contentsTime, ($contentsTime | Select-String -Pattern "#Edited on:").Line)
            $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
            $contentsTime[$index_edit_time] = "#Edited on: $date"
            Set-Content -Path $newScriptDevPath -Value $contentsTime
        }
    
}


    if ($StatusChechkbox_newSensor_creator -eq "true"){
        Write-Host "Creator:"$newScriptCreator

            $contents = Get-Content $newScriptDevPath
            $index = [array]::IndexOf($contents, ($contents | Select-String -Pattern "#Edited by:").Line)
            if ($index -eq -1) {
                return
                Write-Host "error"
            } else {
                $contents[$index] = "#Edited by: $newScriptCreator"
                Set-Content -Path $newScriptDevPath -Value $contents
            $contentsTime = Get-Content $newScriptDevPath
            $index_edit_time = [array]::IndexOf($contentsTime, ($contentsTime | Select-String -Pattern "#Edited on:").Line)
            $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
            $contentsTime[$index_edit_time] = "#Edited on: $date"
            Set-Content -Path $newScriptDevPath -Value $contentsTime
            }
        
    }

    if ($StatusChechkbox_newSensor_changlog -eq "true"){
        Write-Host "ChangeLog:"$newScriptChanglog
        $selectedScript = $var_lb_selected_scripts.SelectedItem

            $contents = Get-Content $newScriptDevPath
            $index = [array]::IndexOf($contents, ($contents | Select-String -Pattern "#Changelog:").Line)
            if ($index -eq -1) {
                return
                Write-Host "error"
            } else {
                $contents[$index] = "#Changelog: $newScriptChanglog"
                Set-Content -Path $newScriptDevPath -Value $contents
            $contentsTime = Get-Content $newScriptDevPath
            $index_edit_time = [array]::IndexOf($contentsTime, ($contentsTime | Select-String -Pattern "#Edited on:").Line)
            $date = Get-Date -Format "dd/MM/yy - hh:mm:ss"
            $contentsTime[$index_edit_time] = "#Edited on: $date"
            Set-Content -Path $newScriptDevPath -Value $contentsTime
            }
        
    }

    $userName = $env:UserName
    $contentsCreator = Get-Content $newScriptDevPath
    $index_edit_Creator = [array]::IndexOf($contentsCreator, ($contentsCreator | Select-String -Pattern "#Created by:").Line)
    $contentsCreator[$index_edit_Creator] = "#Created by: $userName"
    Set-Content -Path $newScriptDevPath -Value $contentsCreator
    
    $contentsCreateTime = Get-Content $newScriptDevPath
    $index_edit_Createtime = [array]::IndexOf($contentsCreateTime, ($contentsCreateTime | Select-String -Pattern "#Created on:").Line)
    $CreateDate = Get-Date -Format "dd/MM/yy - hh:mm:ss"
    $contentsCreateTime[$index_edit_Createtime] = "#Created on: $CreateDate"
    Set-Content -Path $newScriptDevPath -Value $contentsCreateTime
        



})



#Checkbox_: "cb_new_sensor_name"
$var_cb_new_sensor_changlog.Add_Unchecked({
  $var_bt_add_desc_changes.IsEnabled = $false
})

#Checkbox_: "cb_new_sensor_name"
$var_cb_new_sensor_changlog.Add_checked({
  $var_bt_add_desc_changes.IsEnabled = $true
  $window.TopMost = $false
})


#Checkbox_: "cb_new_sensor_name"
$var_cb_new_sensor_name.Add_Unchecked({
  $var_btn_new_sensor.IsEnabled = $false
})

#Checkbox_: "cb_new_sensor_name"
$var_cb_new_sensor_name.Add_checked({
  $var_btn_new_sensor.IsEnabled = $true
  $window.TopMost = $false
})



#Checkbox_: "Clear dev folder"
$var_cb_clear_dev_folder.Add_Unchecked({
  $var_bt_clear_dev_folder.IsEnabled = $false
})

#Checkbox_: "Clear dev folder"
$var_cb_clear_dev_folder.Add_checked({
  $var_bt_clear_dev_folder.IsEnabled = $true
  $window.TopMost = $false
})


#Button: "clear dev folder"
$var_bt_clear_dev_folder.Add_Click({
$window.TopMost = $false

$devFileCount = Get-ChildItem $devPath | Measure-Object
if ($devFileCount.count -ne 0) {
    $date = Get-Date -Format "dd-MM-yy-hh-mm-ss"
    $devSaveFolder = $oldPath + "\dev_save_" + $date
    New-Item -Path $devSaveFolder -ItemType directory

    robocopy $devPath $devSaveFolder /mir
    Get-ChildItem $devPath | ForEach-Object {Remove-Item $_.FullName -Force}
    Write-Host "Dev Folder is clear now | Backup moved to OLD folder" -ForegroundColor Yellow
    explorer.exe $devSaveFolder

} else {
    Write-Host "The folder is empty. nothing happend"
    [System.Windows.MessageBox]::Show('The folder is empty. nothing happend!', 'PRTG Script Manager Message', 'ok')
}

    $var_tb_script_info.Text = ""
    $var_tb_selected_script.Text = "none"
    $var_bt_clear_dev_folder.IsEnabled = $false
    $var_cb_clear_dev_folder.IsChecked = $False
    $var_bt_run_in_ise.isEnabled = $false


})







#Your Code end
##############################################################################################################

  # Show the XAML GUI
  if (!$window.IsLoaded) {
    $window.TopMost = $true
    $Null = $window.ShowDialog()
  }
