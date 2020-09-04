#schtasks /query /tn "AvisMaintenance" /xml >> "c:\temp\AvisMaintenancesTask.xml"
Function DeleteTaskAvis
{
    & SchTasks.exe /DELETE /TN "\AvisMaintenance1" /f
    & SchTasks.exe /DELETE /TN "\AvisMaintenance2" /f
}

Function CreateTaskAvis
{
    schtasks /create /tn "AvisMaintenance1" /xml "c:\temp\AvisMaintenancesTask.xml" 
    schtasks /change /TN AvisMaintenance1 /RU "NT AUTHORITY\INTERACTIVE"
    schtasks /create /tn "AvisMaintenance2" /xml "c:\temp\AvisMaintenances2Task.xml" 
    schtasks /change /TN AvisMaintenance2 /RU "NT AUTHORITY\INTERACTIVE"
}

function Download2
{param($Source,$Ouput)

    $WebClient = New-Object System.Net.WebClient
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
    $WebClient.DownloadFile($Source,$Ouput)
    $WebClient.Dispose()

}

function DownloadGithub
{
    $files = @()
    Download2 -Source "https://github.com/SebasHilotech/OfficeUninstall/raw/master/AvisMaintenace1.ps1" -Ouput "c:\temp\AvisMaintenace1.ps1"
    $file = New-Object System.Object
    $Name = "AvisMaintenace1.ps1"
    $file | Add-Member -type NoteProperty -name NAME -Value "$Name"
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\AvisMaintenace1.ps1"
    $files += $file

    Download2 -Source "https://github.com/SebasHilotech/OfficeUninstall/raw/master/AvisMaintenace2.ps1" -Ouput "c:\temp\AvisMaintenace2.ps1"
    $file = New-Object System.Object
    $Name = "AvisMaintenace2.ps1"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\AvisMaintenace2.ps1"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/AvisMaintenances2Task.xml" -Ouput "c:\temp\AvisMaintenances2Task.xml"
    $file = New-Object System.Object
    $Name = "AvisMaintenances2Task.xml"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\AvisMaintenances2Task.xml"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/AvisMaintenancesTask.xml" -Ouput "c:\temp\AvisMaintenancesTask.xml"
    $file = New-Object System.Object
    $Name = "AvisMaintenancesTask.xml"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\AvisMaintenancesTask.xml"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/runavis1.bat" -Ouput "c:\temp\runavis1.bat"
    $file = New-Object System.Object
    $Name = "runavis1.bat"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\runavis1.bat"
    $files += $file

    Download2 -Source "https://raw.githubusercontent.com/SebasHilotech/OfficeUninstall/master/runavis2.bat" -Ouput "c:\temp\runavis2.bat"
    $file = New-Object System.Object
    $Name = "runavis2.bat"
    $file | Add-Member -type NoteProperty -name NAME -Value $Name
    $file | Add-Member -type NoteProperty -name FULLPATH -Value "c:\temp\runavis2.bat"
    $files += $file
    return $files
}

$Files = DownloadGithub
CreateTaskAvis
schtasks /change /TN AvisMaintenance1 /RU "NT AUTHORITY\INTERACTIVE"
schtasks /Run /TN "AvisMaintenance1"
schtasks /Run /TN "AvisMaintenance2"
Start-Sleep -Seconds 10
#DeleteTaskAvis

foreach($file in $files)
{
    $path = $file.FULLPATH 
    Remove-item $path
}