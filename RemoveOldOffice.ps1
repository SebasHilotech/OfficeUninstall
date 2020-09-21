$WorkDirectory = "c:\temp"

$ErrorActionPreference= 'silentlycontinue'
function setRMM
{param($RMM)
    if($RMM -ne "")
    {
        if(!(Test-Path "C:\temp\RMM.TXT" )){ New-Item -ItemType File -Path "C:\temp\" -Name "RMM.TXT" -Force }
    }
    $TASK  = Get-Content "C:\temp\RMM.TXT"
    if($null -eq  $TASK )
    {
        #Set-Content -Value $RMM -Path "C:\temp\RMM.TXT"
        return $true
    }
    elseif ($TASK -eq $RMM)
    {
        return $false
    }
    else
    {
        Set-Content -Value $RMM -Path "C:\temp\RMM.TXT"
        return $true
    }
}

function getExe
{param($UninstallString)

    $index1 = $UninstallString.IndexOf('"')
    $index2 = $UninstallString.IndexOf('"',$index1 +1)
    $exe = $UninstallString.Substring($index1+1,$index2-1)
    return $exe
}
function getExeParam
{param($UninstallString)

    $index1 = $UninstallString.IndexOf('"')
    $index2 = $UninstallString.IndexOf('"',$index1 +1)
    $exe = $UninstallString.Substring($index1,$index2+1)
    $ExeParam = $UninstallString.Replace($exe,"") 
    return $ExeParam
}

function GetBasicXML
{param($type)
    $BasicXML = @"
<Configuration Product="$type">
    <Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
    <Setting Id="SETUP_REBOOT" Value="Never" />
</Configuration>
"@
return $BasicXML
}
function GetUninstall365XML
{param($productid)
$XML =    @"
<Configuration>
<Remove>
<Product ID="$productid">
<Property Name="FORCEAPPSHUTDOWN" Value="True" />
<Property Name="PinIconsToTaskbar" Value="False"/>
</Product>
</Remove>
<Property Name="FORCEAPPSHUTDOWN" Value="True" />
<Logging Level="Standard" Path="c:\temp\officelog.log" />
<Display Level="None" AcceptEULA="TRUE" CompletionNotice="no" SuppressModal="yes" />
<Updates Enabled="TRUE" Branch="Current"/>
</Configuration>
"@
return $XML
}
function GetOffice365Productid
{
    $OfficeObject = GetOfficeVersion
    $index1 = $OfficeObject.param.IndexOf("productstoremove=")
    $index2 = $OfficeObject.param.IndexOf(" ",$OfficeObject.param.IndexOf("productstoremove="))
    $index3 = $index2 - $index1
    $ProductidNotClean = ($OfficeObject.param.Substring($index1,$index3)).Replace("productstoremove=","")
    $index4 = $ProductidNotClean.IndexOf(".")
    $Productid = $ProductidNotClean.Substring(0,$index4)
    return $Productid
}
function GetUninstallXML
{param($version)
    if($version -match "Microsoft 365")
    {
        $Productid = GetOffice365Productid
        $XML = GetUninstall365XML -productid $Productid
    }
    elseif($version -match "Standard"){
        $XML =       GetBasicXML -type "Standard"
    }
    elseif($version -match "Professional Plus")
    {
        $XML = GetBasicXML -type "ProPlus"
    }elseif($version -match "Small Business")
    {
        $XML = GetBasicXML -type "SMALLBUSINESS"
    }
    $WorkFolder = "C:\temp"
    $UninstallXML = "uninstallOldOffice.xml"
    $UninstallXMLFullName = $WorkFolder + "\" + $UninstallXML
    if(!(Test-Path $UninstallXMLFullName ))
    {
        New-Item -ItemType File -Path $WorkDirectory -Name $UninstallXML  -Force | Out-Null
    }
    Set-Content -Value $XML -Path $UninstallXMLFullName | Out-Null
    return $UninstallXMLFullName
}
function getExeParamSilent
{param($exe,$Version)
    if($exe -like "*OfficeClickToRun*")
    {
        $ParamSilent = " DisplayLevel=False"
    }
    elseif($exe -like "*setup.exe*")
    {
        $UninstallXMLPath = GetUninstallXML -version $Version
        $ParamSilent = " /config " + $UninstallXMLPath
    }
    return $ParamSilent
}
function GetOfficeVersion
{
    $List32 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    $List64 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    $List = $List32 + $List64
    #Check for MS Office
    $OfficeVersion = @()
    foreach($program in $List)
    {
        if($program.DisplayName -like "Microsoft Office*")
        {
            if($program.UninstallString -match "MsiExec"){}else{$OfficeVersion += $program}
        }
    }
    $ListOfficeObject = @()
    if($OfficeVersion.Count -gt 1){
        foreach($office in $OfficeVersion)
        {
            $OfficeObject = New-Object System.Object
            $UninstallString = $office.UninstallString
            $exe = getExe -uninstallString $UninstallString
            $exeParam = getExeParam -uninstallString $UninstallString
            $name = $office.DisplayName
            $exeParamSilent = getExeParamSilent -exe $exe -UninstallString $UninstallString -Version $name
            $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
            $OfficeObject | Add-Member -type NoteProperty -name "exe" -Value "$exe"
            $OfficeObject | Add-Member -type NoteProperty -name "param" -Value "$exeParam"
            $OfficeObject | Add-Member -type NoteProperty -name "silent" -Value "$exeParamSilent"
            $ListOfficeObject += $OfficeObject
        }
    }
    else
    {
        $OfficeObject = New-Object System.Object
        $name = $OfficeVersion.DisplayName
        $UninstallString = $OfficeVersion.UninstallString
        $name = $OfficeVersion.DisplayName
        $exe = getExe -uninstallString $UninstallString
        $exeParam = getExeParam -uninstallString $UninstallString
        $exeParamSilent = getExeParamSilent -exe $exe -UninstallString $UninstallString -Version $name
        $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
        $OfficeObject | Add-Member -type NoteProperty -name "exe" -Value "$exe"
        $OfficeObject | Add-Member -type NoteProperty -name "param" -Value "$exeParam"
        $OfficeObject | Add-Member -type NoteProperty -name "silent" -Value "$exeParamSilent"
        $ListOfficeObject += $OfficeObject
    }
    return $ListOfficeObject
}
function UninstallOffice
{param($OfficeObject)
    if($OfficeObject.count -gt 1){$OfficeObject = $OfficeObject[0]}

    if($OfficeObject.name -match "365")
    {
        $Exe = $OfficeObject.exe
        $arguments = $OfficeObject.param + " " + $OfficeObject.silent
        start-process $Exe -Args $arguments -Wait -NoNewWindow
    }
    else
    {
        $Exe = $OfficeObject.exe
        $arguments = $OfficeObject.param + " " + $OfficeObject.silent
        start-process $Exe -Args $arguments -Wait -NoNewWindow
    }

    $OfficeObject = CheckIfOfficeStillInstalled
    if($OfficeObject -ne $false){$state = "multi"}else{$state = "done"}
    return $state
}
function InstallO365
{
    $Exe = "C:\temp\setup.exe"
    $XML = "C:\temp\XML\Office365_FR_FR32.xml"
    $arguments = " /configure " + $XML
    start-process $Exe -args $arguments -Wait -NoNewWindow
}
function CheckIfOffice365Installed
{
    $List32 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    $List64 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    $List = $List32 + $List64
    #Check for MS Office
    $OfficeVersion = @()
    foreach($program in $List)
    {
        if($program.DisplayName -like "Microsoft 365*")
        {
            if($program.UninstallString -match "MsiExec"){}else{$OfficeVersion += $program}
        }
        if($program.DisplayName -like "Microsoft Office*365*")
        {
            if($program.UninstallString -match "MsiExec"){}else{$OfficeVersion += $program}
        }
    }
    $ListOfficeObject = @()
    if($OfficeVersion.Count -gt 1){
        foreach($office in $OfficeVersion)
        {
            $OfficeObject = New-Object System.Object
            $name = $office.DisplayName
            $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
            $ListOfficeObject += $OfficeObject
        }
    }
    elseif($OfficeVersion.Count -eq 1)
    {
        $OfficeObject = New-Object System.Object
        $name = $OfficeVersion.DisplayName
        $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
        $ListOfficeObject += $OfficeObject
    }
    else
    {
        $ListOfficeObject = $false
    }
    return $ListOfficeObject
}

function CheckIfOfficeStillInstalled
{ 
    $Office32 = CheckIfOfficeStillInstalled32
    
    if($Office32 -eq $false )
    {
        $Office64 = CheckIfOfficeStillInstalled64
        if($Office64 -eq $false)
        {
            $OfficeObject = $false
        }
        else
        {
            $OfficeObject =  $Office64
            New-Item -Path c:\temp\ -Name "old3264.txt" -Value "64" -Force | Out-Null
        }
    }
    else
    {
        $OfficeObject =  $Office32
        New-Item -Path c:\temp\ -Name "old3264.txt" -Value "32" -Force | Out-Null
    }
    return $OfficeObject
}

function CheckIfOfficeStillInstalled64
{
    #$List32 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    $List32 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    $List = $List32 #+ $List64
    #Check for MS Office
    $OfficeVersion = @()
    foreach($program in $List)
    {
        if($program.DisplayName -like "Microsoft Office*")
        {
            if(($program.UninstallString -match "MsiExec") -or ($program.DisplayName -like "*Language Pack*") -or ($program.DisplayName -like "*viso*") -or ($program.DisplayName -like "*sharepoint*")  -or ($program.DisplayName -like "*project*")){}else{$OfficeVersion += $program}
        }
    }
    $ListOfficeObject = @()
    if($OfficeVersion.Count -gt 1){
        foreach($office in $OfficeVersion)
        {
            $OfficeObject = New-Object System.Object
            $UninstallString = $office.UninstallString
            $exe = getExe -uninstallString $UninstallString
            $exeParam = getExeParam -uninstallString $UninstallString
            $name = $office.DisplayName
            $exeParamSilent = getExeParamSilent -exe $exe -UninstallString $UninstallString -Version $name
            $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
            $OfficeObject | Add-Member -type NoteProperty -name "exe" -Value "$exe"
            $OfficeObject | Add-Member -type NoteProperty -name "param" -Value "$exeParam"
            $OfficeObject | Add-Member -type NoteProperty -name "silent" -Value "$exeParamSilent"
            $OfficeObject | Add-Member -type NoteProperty -name "3264" -Value "64"
            $ListOfficeObject += $OfficeObject
        }
    }
    elseif($OfficeVersion.Count -eq 1)
    {
        $OfficeObject = New-Object System.Object
        $name = $OfficeVersion.DisplayName
        $UninstallString = $OfficeVersion.UninstallString
        $name = $OfficeVersion.DisplayName
        $exe = getExe -uninstallString $UninstallString
        $exeParam = getExeParam -uninstallString $UninstallString
        $exeParamSilent = getExeParamSilent -exe $exe -UninstallString $UninstallString -Version $name

        $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
        $OfficeObject | Add-Member -type NoteProperty -name "exe" -Value "$exe"
        $OfficeObject | Add-Member -type NoteProperty -name "param" -Value "$exeParam"
        $OfficeObject | Add-Member -type NoteProperty -name "silent" -Value "$exeParamSilent"
        $OfficeObject | Add-Member -type NoteProperty -name "3264" -Value "64"
        $ListOfficeObject += $OfficeObject
    }
    else
    {
        $ListOfficeObject = $false
    }
    return $ListOfficeObject
    #$ListOfficeObject64bit
}

function CheckIfOfficeStillInstalled32
{
    $List64 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    #$List32 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, UninstallString | Where-Object {$_.DisplayName -like "*microsoft*"}
    $List = $List64
    #Check for MS Office
    $OfficeVersion = @()
    foreach($program in $List)
    {
        if($program.DisplayName -like "Microsoft Office*")
        {
            if(($program.UninstallString -match "MsiExec") -or ($program.DisplayName -like "*Language Pack*") -or ($program.DisplayName -like "*viso*") -or ($program.DisplayName -like "*sharepoint*")  -or ($program.DisplayName -like "*project*")){}else{$OfficeVersion += $program}
        }
    }
    $ListOfficeObject = @()
    if($OfficeVersion.Count -gt 1){
        foreach($office in $OfficeVersion)
        {
            $OfficeObject = New-Object System.Object
            $UninstallString = $office.UninstallString
            $exe = getExe -uninstallString $UninstallString
            $exeParam = getExeParam -uninstallString $UninstallString
            $name = $office.DisplayName
            $exeParamSilent = getExeParamSilent -exe $exe -UninstallString $UninstallString -Version $name
            $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
            $OfficeObject | Add-Member -type NoteProperty -name "exe" -Value "$exe"
            $OfficeObject | Add-Member -type NoteProperty -name "param" -Value "$exeParam"
            $OfficeObject | Add-Member -type NoteProperty -name "silent" -Value "$exeParamSilent"
            $OfficeObject | Add-Member -type NoteProperty -name "3264" -Value "32"
            $ListOfficeObject += $OfficeObject
        }
    }
    elseif($OfficeVersion.Count -eq 1)
    {
        $OfficeObject = New-Object System.Object
        $name = $OfficeVersion.DisplayName
        $UninstallString = $OfficeVersion.UninstallString
        $name = $OfficeVersion.DisplayName
        $exe = getExe -uninstallString $UninstallString
        $exeParam = getExeParam -uninstallString $UninstallString
        $exeParamSilent = getExeParamSilent -exe $exe -UninstallString $UninstallString -Version $name

        $OfficeObject | Add-Member -type NoteProperty -name "name" -Value "$name"
        $OfficeObject | Add-Member -type NoteProperty -name "exe" -Value "$exe"
        $OfficeObject | Add-Member -type NoteProperty -name "param" -Value "$exeParam"
        $OfficeObject | Add-Member -type NoteProperty -name "silent" -Value "$exeParamSilent"
        $OfficeObject | Add-Member -type NoteProperty -name "3264" -Value "32"
        $ListOfficeObject += $OfficeObject
    }
    else
    {
        $ListOfficeObject = $false
    }
    return $ListOfficeObject
}

function GetLocalUser
{
    $result=@()
    $computer =1
    If ($Computer)
    {
        quser | Select-Object -Skip 1 | Foreach-Object {
                    $b=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
                    If ($b[2] -like 'Disc*')
                    {
                        $array= ([ordered]@{
                            'User' = $b[0]
                            'ID'   = $b[2]
                            'Date' = $b[4]
                            'Time' = $b[5..6] -join ' '})

                        $result+=New-Object -TypeName PSCustomObject -Propertyrty $array
                    }
                    else
                    {
                        $array= ([ordered]@{
                        'User' = $b[0]
                        'ID'   = $b[2]
                        'Date' = $b[5]
                        'Time' = $b[6..7] -join ' '
                    })
                $result+=New-Object -TypeName PSCustomObject -Property $array
            }
        }
    }
    return $result
}
function Expand-ZIPFile($file, $destination)
{
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
}
function Download
{param($id,$DownloadPath)
    $WebClient = New-Object System.Net.WebClient
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
   # $url = "https://github.com/SebasHilotech/Production/blob/master/Remove-PreviousOfficeInstalls.zip"
    $url = "https://drive.google.com/uc?id=" + $id + "&export=download"
    $DownloadPath = "C:\temp\Remove-PreviousOfficeInstalls.zip"
    $WebClient.DownloadFile($url,$DownloadPath)
    $WebClient.Dispose()
}
function GetFile
{param($Project)

    $ListDownloadPath = $Project.LISTDOWNLOAD
    $workdirectory    = $Project.WORKFOLDER
    $XmlDirectory     = $Project.XMLDIRECTORY
    $listDownload     = import-csv $ListDownloadPath
    $files = @()
    foreach($item in $listDownload)
    {
        $id = $item.Key
        if($item.Install -eq "WORK")
        {
            $DownloadPath = $workdirectory + "\" + $item.FullName
        }
        elseif($item.Install -eq "XML")
        {
            $DownloadPath = $XmlDirectory + "\" + $item.FullName
        }
        if(!(Test-Path $DownloadPath))
        {
            $id = $item.Key
            Download -id $id -DownloadPath $DownloadPath
        }
        $file = New-Object System.Object
        $Name = $item.Item
        $file | Add-Member -type NoteProperty -name NAME -Value "$Name"
        $file | Add-Member -type NoteProperty -name FULLPATH -Value "$DownloadPath"
        $files += $file
    }
    return $files
}
function IsUserInList
{param($ListFile,[switch]$test)
    $CSVLISTUSER = $ListFile | where-Object {$_.NAME -eq "ListUserInstallation"}
    $LISTUSER = import-csv $CSVLISTUSER.FULLPATH -Encoding utf8
    $CurrentUser = GetLocalUser
    #Test only
    if($test -eq $true)
    {
        $CurrentUser = "vlare"
    }
    if($LISTUSER | Where-Object {$_.User -eq $CurrentUser})
    {
        $InList = $true
    }else
    {
        $InList = $false
    }
    return $InList
}
function incrementReboot
{param([int]$maxReboot=6,$Project)
    $WorkFolder = $Project.WORKFOLDER
    $rebootFile = "boot.csv"
    $rebootFullPath = "$WorkFolder\" + $rebootFile
    $arrayReboot = @()
    if(!(Test-Path $rebootFullPath)){
        $allowReboot = $true
        $reboot = New-Object System.Object
        $reboot | Add-Member -type NoteProperty -name RebootIncrement -Value 0
        $reboot | Add-Member -type NoteProperty -name Time -Value $(get-date -Format "yyyy/mm/dd HH:mm:ss")
        $reboot | Add-Member -type NoteProperty -name Output -Value "Not set"
        $reboot | Add-Member -type NoteProperty -name Reboot -Value $allowReboot
        $arrayReboot += $reboot
        $arrayReboot | export-csv -Path $rebootFullPath | Out-Null
    }
    else
    {
        $csvReboot = import-csv -Path $rebootFullPath -Encoding utf8

        if($csvReboot.count -lt $maxReboot)
        {
            $arrayReboot += $csvReboot
            $count = $arrayReboot.count
            $rebootCount = $count++
            $allowReboot = $true
            $reboot = New-Object System.Object
            $reboot | Add-Member -type NoteProperty -name RebootIncrement -Value $rebootCount
            $reboot | Add-Member -type NoteProperty -name Time -Value $(get-date -Format "yyyy/mm/dd HH:mm:ss")
            $reboot | Add-Member -type NoteProperty -name Output -Value "Not set"
            $reboot | Add-Member -type NoteProperty -name Reboot -Value $allowReboot
            $arrayReboot += $reboot
            $arrayReboot | export-csv -Path $rebootFullPath | Out-Null
        }
        else
        {
            $csvReboot = import-csv -Path $rebootFullPath -Encoding utf8
            $arrayReboot += $csvReboot
            $count = $arrayReboot.count
            $rebootCount = $count++
            $allowReboot = $false
            $reboot = New-Object System.Object
            $reboot | Add-Member -type NoteProperty -name RebootIncrement -Value $rebootCount
            $reboot | Add-Member -type NoteProperty -name Time -Value $(get-date -Format "yyyy/mm/dd HH:mm:ss")
            $reboot | Add-Member -type NoteProperty -name Output -Value "error, tried to boot to often"
            $reboot | Add-Member -type NoteProperty -name Reboot -Value $allowReboot
            $arrayReboot += $reboot
            $arrayReboot | export-csv -Path $rebootFullPath
        }
    }
    return $allowReboot
}
function Reboot
{

    & shutdown.exe -t 0 -r -f
    Start-Sleep -Seconds 90

}
function GetStep
{param($Project)

    $WorkFolder = $Project.WORKFOLDER
    $stepfile = "step.ini"
    $stepfileFullName = $WorkFolder + "\" + $stepfile

    if(!(Test-Path $stepfileFullName ))
    {
        New-Item -ItemType File -Path $WorkDirectory -Name $stepfile  -Force | Out-Null
        Set-Content -Value "0" -Path $stepfileFullName | Out-Null
    }
    $step = Get-Content -Path $stepfileFullName -Encoding utf8
    if($null -eq $step)
    {
        Set-Content -Value "0" -Path $stepfileFullName | Out-Null
        $step = Get-Content -Path $stepfileFullName -Encoding utf8
    }
    return $step
}
function IsWindows10
{
    if([system.Environment]::Osversion.Version.Major -eq 10)
    {
        return $true
    }
    else
    {
        return $false
    }
}
function CreateTaskMigration
{
    if(IsWindows10)
    {
        $Trigger =   New-ScheduledTaskTrigger -AtStartup
        $Trigger.Delay = 'PT2M'
        $User= "NT AUTHORITY\SYSTEM" # Specify the account to run the script
        $setting = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries
        $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument " -executionPolicy Unrestricted  -file ""C:\temp\deploiementOffice365.ps1""" # Specify what program to run and with its parameters
        Register-ScheduledTask -TaskName "Office365Migration" -Trigger $Trigger -User $User -Action $Action -RunLevel Highest -Settings $setting  -Force # Specify the name of the task
    }
    else
    {
        CreateTaskPS2
    }
}
function CreateTaskPS2
{
    $TaskName = "Office365Migration"
    $TaskDescription = "Install PS5_1-UninstallOffice&InstallOffice365"
    $TaskCommand = "c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe"
    $TaskScript = "C:\temp\deploiementOffice365.ps1"
    $TaskArg = "-WindowStyle Hidden -NonInteractive -Executionpolicy unrestricted -file $TaskScript"
    $TaskStartTime = [datetime]::Now.AddMinutes(1)
    $service = new-object -ComObject("Schedule.Service")
    $service.Connect()
    $rootFolder = $service.GetFolder("\")
    $TaskDefinition = $service.NewTask(0)
    $TaskDefinition.RegistrationInfo.Description = "$TaskDescription"
    $TaskDefinition.Settings.Enabled = $true
    $TaskDefinition.Settings.AllowDemandStart = $true
    $TaskDefinition.Settings.DisallowStartIfOnBatteries = $false
    $TaskDefinition.Principal.RunLevel = 1
    $TaskDefinition.Settings.StopIfGoingOnBatteries = $false
    $triggers = $TaskDefinition.Triggers
    #http://msdn.microsoft.com/en-us/library/windows/desktop/aa383915(v=vs.85).aspx
    $trigger = $triggers.Create(8)
    $trigger.StartBoundary = $TaskStartTime.ToString("yyyy-MM-dd'T'HH:mm:ss")
    $trigger.Enabled = $true
    $trigger.Delay = "PT2M"
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa381841(v=vs.85).aspx
    $Action = $TaskDefinition.Actions.Create(0)
    $action.Path = "$TaskCommand"
    $action.Arguments = "$TaskArg"
    #http://msdn.microsoft.com/en-us/library/windows/desktop/aa381365(v=vs.85).aspx
    $rootFolder.RegisterTaskDefinition("$TaskName",$TaskDefinition,6,"System",$null,5)
}
Function DeleteTaskMigration
{
    if(IsWindows10)
    {
        Unregister-ScheduledTask  -TaskPath "\" -TaskName "Office365Migration" -Confirm:$false
    }
    else
    {
        & SchTasks.exe /DELETE /TN "\Office365Migration" /f
    }
}
function IncrementStep
{param($Project)
    $WorkFolder = $Project.WORKFOLDER
    $stepfile = "step.ini"
    $stepfileFullName = $WorkFolder + "\" + $stepfile
    if(!(Test-Path $stepfileFullName ))
    {
        New-Item -ItemType File -Path $WorkDirectory -Name $stepfile  -Force | Out-Null
    }
    $step = Get-Content -Path $stepfileFullName -Encoding utf8
    $stepInt = $step -as [int]
    $stepInt++
    Set-Content -Value $stepInt -Path $stepfileFullName
}

function IsPowershell51
{
    if($PSVersionTable.PSVersion.Major -lt 5){return $false}
    else{return $true}
}
function Powershell51Install
{
    $WebClient = New-Object System.Net.WebClient
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
    $WebClient.DownloadFile("http://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7AndW2K8R2-KB3191566-x64.zip","C:\temp\Win7AndW2K8R2-KB3191566-x64.zip")
    Expand-ZIPFile -file "C:\temp\Win7AndW2K8R2-KB3191566-x64.zip" -destination "c:\temp\"
    Remove-Item -Path "C:\temp\Win7AndW2K8R2-KB3191566-x64.zip"
    Set-Location -Path "C:\temp"
    . .\Install-WMF5.1.ps1 -AcceptEULA -AllowRestart -Confirm:$false
    # it will skip and continue if it's not there.
    Start-Sleep -Seconds 90
    Reboot
}

function OPUninstall
{
    #Download -id "1ClipcaQFlHLkZjc7Wj0zS8-8LAbk41xn" -DownloadPath "c:\temp\Remove-PreviousOfficeInstalls.zip"
    $WebClient = New-Object System.Net.WebClient
    $url = "https://github.com/SebasHilotech/OfficeUninstall/raw/master/Remove-PreviousOfficeInstalls.zip"
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
    $WebClient.DownloadFile($url,"c:\temp\Remove-PreviousOfficeInstalls.zip")

    #expand-archive -path "c:\temp\Remove-PreviousOfficeInstalls.zip" -DestinationPath "C:\temp\"
    Expand-ZIPFile -file "c:\temp\Remove-PreviousOfficeInstalls.zip" -destination "C:\temp\"
    Remove-Item "c:\temp\Remove-PreviousOfficeInstalls.zip"
    Set-Location -path  "C:\temp\Remove-PreviousOfficeInstalls"

    . .\Remove-PreviousOfficeInstalls.ps1 -ProductsToRemove MainOfficeProduct -Quiet:$true -Force:$true -Remove2016Installs:$true

    Remove-PreviousOfficeInstalls

    set-Location -Path "C:\temp"
}
function CallStep
{param($Project,$ListFiles)
$step = GetStep -Project $Project
switch($step)
{
    "0"
    {
        $ListUser = GetLocalUser
        $ID = $ListUser.ID
        logoff $ID
        CreateTaskMigration
        if(!(IsPowershell51))
        {
            Powershell51Install
        }
        IncrementStep -Project $Project
    }
    "1"
    {
        $ListUser = GetLocalUser
        $ID = $ListUser.ID
        logoff $ID
        $OfficeObject = CheckIfOfficeStillInstalled
        if($OfficeObject -ne $false)
        {
            if($OfficeObject.count -gt 1)
            {
                foreach($office in $OfficeObject )
                {
                    $resultUninstall = UninstallOffice -OfficeObject $office 
                }
            }
            else
            {
                $resultUninstall =  UninstallOffice -OfficeObject $OfficeObject
            }
        }
        $OfficeObject = CheckIfOfficeStillInstalled
        if($OfficeObject -ne $false)
        {
            OPUninstall
        }
        IncrementStep -Project $Project
    }
    "2"
    {
        $OfficeObject = CheckIfOfficeStillInstalled
        if($OfficeObject -eq $false)
        {
            $resultUninstall = UninstallOffice -OfficeObject $OfficeObject
        }
        #Install Microsoft Office 365
        InstallO365 -ListFiles $ListFiles
        IncrementStep -Project $Project
        Reboot
    }
    "3"
    {
        $ListUser = GetLocalUser
        $ID = $ListUser.ID
        logoff $ID
        $result  = CheckIfOffice365Installed
        if($result -eq $false)
        {
            InstallO365 -ListFiles $ListFiles
        }
        DeleteTaskMigration
        DeleteStuff -ListFiles $ListFiles
        IncrementStep -Project $Project
        Reboot
    }
    Default
    {
        # if something else, delete the task
        DeleteTaskMigration
    }
  }
}
function DeleteStuff
{param($ListFiles)
    foreach($file in $ListFiles)
    {
        $item = $file.FULLPATH
        if(test-path $item){Remove-Item -Path $item}
    }
    if(Test-Path -Path "C:\temp\Remove-PreviousOfficeInstalls")
    {
        Get-ChildItem "C:\temp\Remove-PreviousOfficeInstalls" | Remove-Item
        Remove-Item "C:\temp\Remove-PreviousOfficeInstalls"
    }
    if(Test-Path -Path "C:\temp\ListDownloadBraun.csv")
    {
        Remove-Item "C:\temp\ListDownloadBraun.csv"
    }
    if(Test-Path -Path "C:\temp\XML")
    {
        Remove-Item "C:\temp\XML"
    }
    $listIconsPath = Get-ChildItem "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 20*"
    foreach ( $item in $listIconsPath )
    {
        if($item.Extension -eq "lnk")
        {
            Remove-Item $Item.FullName
        }
    }
}

function Download2
{param($Source,$Ouput)

    $WebClient = New-Object System.Net.WebClient
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
    $WebClient.DownloadFile($Source,$Ouput)
    $WebClient.Dispose()

}
function downlaodGitHub
{
    $files = @()

    Download2 -Source "https://github.com/SebasHilotech/OfficeUninstall/raw/master/setup.exe" -Ouput "c:\temp\setup.exe"
    $file = New-Object System.Object
    $Name = "ODT"
    $file | Add-Member -type NoteProperty -name NAME -Value "$Name"
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\setup.exe"
    $files += $file

    Download2 -Source "https://github.com/SebasHilotech/OfficeUninstall/raw/master/OfficeBusinessRetail_Install_EN_US.xml" -Ouput "c:\temp\XML\OfficeBusinessRetail_Install_EN_US.xml"
    $file = New-Object System.Object
    $Name = "OfficeBusinessRetail_Install_EN_US"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\XML\OfficeBusinessRetail_Install_EN_US.xml"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/OfficeProPlus_Install_FR_fr.xml" -Ouput "c:\temp\XML\OfficeProPlus_Install_FR_fr.xml"
    $file = New-Object System.Object
    $Name = "OfficeProPlus_Install_EN_US"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\XML\OfficeProPlus_Install_EN_US.xml"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/deploiementOffice365.ps1" -Ouput "c:\temp\deploiementOffice365.ps1"
    $file = New-Object System.Object
    $Name = "deploiementOffice365"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\deploiementOffice365.ps1"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/Office365_EN_EN_FR_FR64.xml" -Ouput "c:\temp\Office365_FR_FR32.xml"
    $file = New-Object System.Object
    $Name = "deploiementOffice365"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\Office365_FR_FR32.xml"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/Office365_EN_EN_FR_FR32.xml" -Ouput "c:\temp\XML\Office365_FR_FR32.xml"
    $file = New-Object System.Object
    $Name = "deploiementOffice365"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\XML\Office365_FR_FR32.xml"
    $files += $file
    return $files
}
#endregion


#region set variable, folder and run
$XmlDirectory = "$WorkDirectory\XML"
if(!(Test-Path $XmlDirectory )){ New-Item -ItemType directory -Path $WorkDirectory -Name "XML" -Force }


#Set object for function
$Project = New-Object System.Object
$Project | Add-Member -type NoteProperty -name CSVPROJECT -Value "$projectid"
$Project | Add-Member -type NoteProperty -name WORKFOLDER -Value "$WorkDirectory"
$Project | Add-Member -type NoteProperty -name XMLDIRECTORY -Value "$XmlDirectory"
$Project | Add-Member -type NoteProperty -name LISTDOWNLOAD -Value  "$WorkDirectory\$listDownload"

#$ListFiles = downlaodGitHub

$resultValue  = ""

$ListUser = GetLocalUser
$ID = $ListUser.ID
logoff $ID
$OfficeObject = CheckIfOfficeStillInstalled

if($OfficeObject -ne $false)
{
    if($OfficeObject.count -gt 1)
    {
        foreach($office in $OfficeObject )
        {
            $resultUninstall = UninstallOffice -OfficeObject $office 
        }
    }
    else
    {
        $resultUninstall =  UninstallOffice -OfficeObject $OfficeObject
    }
}
$OfficeObject = CheckIfOfficeStillInstalled
if($OfficeObject -ne $false)
{
    OPUninstall
    $resultValue += "OP"
}

$OfficeObject = CheckIfOfficeStillInstalled
if($OfficeObject -eq $false)
{
    $resultValue += "Done"
    New-Item -Path c:\temp\ -Name "UninstallOffice.txt" -Value $resultValue -Force
}else
{
    $resultValue += "Failed"
    New-Item -Path c:\temp\ -Name "UninstallOffice.txt" -Value $resultValue -Force
}

return $resultValue