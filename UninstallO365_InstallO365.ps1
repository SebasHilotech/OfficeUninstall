#powershell


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
function Download2
{param($Source,$Ouput)

    $WebClient = New-Object System.Net.WebClient
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
    $WebClient.DownloadFile($Source,$Ouput)
    $WebClient.Dispose()

}

if(!(Test-Path "c:\temp\XML" )){ New-Item -ItemType directory -Path "C:\temp" -Name "XML" -Force }

Download2 -Source "https://github.com/SebasHilotech/OfficeUninstall/raw/master/OfficeProPlus_Install_FR_fr.xml" -Ouput "c:\temp\XML\OfficeProPlus_Install_FR_fr.xml"

Download2 -Source "https://github.com/SebasHilotech/OfficeUninstall/raw/master/setup.exe" -Ouput "c:\temp\setup.exe"

$Exe = "C:\temp\setup.exe"

$XML = "c:\temp\XML\OfficeProPlus_Install_FR_fr.xml"

$arguments = " /configure " + $XML 

start-process $Exe -args $arguments -Wait -NoNewWindow
