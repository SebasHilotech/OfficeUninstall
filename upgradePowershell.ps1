function Expand-ZIPFile($file, $destination)
{
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item)
    }
}


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