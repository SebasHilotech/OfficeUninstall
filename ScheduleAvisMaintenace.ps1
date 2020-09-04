#schtasks /query /tn "AvisMaintenance" /xml >> "c:\temp\AvisMaintenancesTask.xml"
Function DeleteTaskAvis
{
    & SchTasks.exe /DELETE /TN "\AvisMaintenance1" /f
    & SchTasks.exe /DELETE /TN "\AvisMaintenance2" /f
}

Function CreateTaskAvis
{
    schtasks /create /tn "AvisMaintenance1" /xml "c:\temp\AvisMaintenancesTask.xml" 
    schtasks /create /tn "AvisMaintenance2" /xml "c:\temp\AvisMaintenances2Task.xml" 
}

#chtasks /Run /TN "AvisMaintenance1"
#schtasks /Run /TN "AvisMaintenance2"
