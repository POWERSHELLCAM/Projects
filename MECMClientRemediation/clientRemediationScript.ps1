
# ######################################################################
#
#     Program         :   ConfigMGrClientHealthCheckup
#     Version         :   1.0
#     Purpose         :   Remediate Config MGr Client
#     Author          :   Shishir Kushawaha
#     Mail Id         :   srktcet@gmail.com
#     Created         :   19-10-2022 - Script creation.
#     Modified        :   24-11-2022 - Added services BITS and MSISERVER
#                                    - Made script more compact and easy to use
#                       
# ######################################################################

$logpath="C:\windows\ccm\logs\ConfigMGrClientHealthCheckup.log"
#Enable PowerShell Execution policy
Try 
{ 
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop 
} Catch 
{
    Write-Output "[$(get-date -format G)] $_ " | Out-File $logpath -Append
}

$servicesList=@('BITS','CCMEXEC','WUAUSERV','MSISERVER')
$deleteFileList=@('C:\Windows\System32\GroupPolicy\Machine\registry.pol','C:\Windows\System32\GroupPolicy\gpt.ini','C:\Windows\smscfg.ini')
$renamefolderList=@('C:\Windows\SoftwareDistribution','C:\Windows\system32\catroot2')
$ScheduleIDMappings = @('{00000000-0000-0000-0000-000000000021}','{00000000-0000-0000-0000-000000000022}','{00000000-0000-0000-0000-000000000026}','{00000000-0000-0000-0000-000000000027}','{00000000-0000-0000-0000-000000000113}','{00000000-0000-0000-0000-000000000114}')

# 1 Stop required services
foreach($service in $servicesList)
{
    Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
    Write-Output "[$(get-date -format G)] Stop $service Service : $?" | Out-File $logpath -Append
}

# 2 Delete required files
foreach($file in $deleteFileList)
{
    Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
    Write-Output "[$(get-date -format G)] Delete $file : $?" | Out-File $logpath -Append
}

# 3 Rename required files
foreach($folder in $renamefolderList)
{
    if(test-path $folder+'_old')
    {
        Remove-Item -Path $folder+'_old' -Recurse -Confirm:$false -Force
    }
    Rename-Item -Path $folder $folder+'_old' -Force -ErrorAction SilentlyContinue
    Write-Output "[$(get-date -format G)] Rename $folder Folder : $?" | Out-File $logpath -Append
}

# 4 Stop required services
foreach($service in $servicesList)
{
    Start-Service -Name $service -Force -ErrorAction SilentlyContinue
    Write-Output "[$(get-date -format G)] Start $service Service : $?" | Out-File $logpath -Append
}

Start-Sleep 120

# 5 Trigger User, machine and update cycles
foreach($id in $ScheduleIDMappings)
{
    [void]([wmiclass] "root\ccm:SMS_Client").TriggerSchedule($id)
    Write-Output "[$(get-date -format G)] Trigger $id : $?" | Out-File $logpath -Append
}




