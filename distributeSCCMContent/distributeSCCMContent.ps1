function writeLog($message)
{
    Write-Output "`n $(get-date -format G)> $message" | out-file $logfile -Append
    Write-Host "$message"
    if($true)
    {
        $speak.Speak($message)
    }
}
function InitializeSCCM  
{
    try 
    {
           # Customizations  
        $initParams = @{}  
        
        # Import the ConfigurationManager.psd1 module   
        if($null -eq (Get-Module ConfigurationManager)) 
        {  
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams   
        }  
        
        # Connect to the site's drive if it is not already present  
        if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) 
        {  
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams  
        }  
        
        # Set the current location to be the site code.  
        Set-Location "$($SiteCode):\" @initParams   
        writeLog "MECM initialized successfully."
    }
    catch 
    {
        writeLog "Failed to initialize MECM. Error : $_"
        exit 1
    }
}  

function intializeSQLConnection
{
    $global:sqlConn = New-Object System.Data.SqlClient.SqlConnection
    $global:sqlConn.ConnectionString = "Server=$dbserver;Integrated Security=true;Initial Catalog=$dbcatalog"
    $global:sqlConn.open()
    $global:sqlcmd = New-Object System.Data.SqlClient.SqlCommand
    $global:sqlcmd.Connection = $sqlConn
    writeLog "DBServer : $dbserver `t Database : $dbcatalog"
    InitializeSCCM
}

function executeSQLQuery
{
    try 
    {
        $list=$null
        $global:sqlcmd.CommandText=$null
        $global:sqlcmd.CommandText = $query
        $adp = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcmd
        $adp.SelectCommand.CommandTimeout=120
        $data = New-Object System.Data.DataSet
        $adp.Fill($data) | Out-Null
        $list=$data.tables
        $data=$adp=$null 
        return $list
        writeLog "Successfully executed query."       
    }
    catch 
    {
        writeLog "Failed to executed query. Error $_" 
        exit 1
    }
}

function checkPackageStatus($packageid)
{
    Writelog "Checking package status of $packageid."
    $global:query="select * from  v_PackageStatusDistPointsSumm where ServerNALPath like '%$distributionPoint%' and PackageID ='$packageid'"
    return executeSQLQuery
}

function checkMissingPackages
{
    $global:query="SELECT DISTINCT PS.PackageID,p.PackageType,n.SourceSize, PSD.Name FROM v_TaskSequencePackageReferences AS TSR
    JOIN v_PackageStatus AS PS ON TSR.RefPackageID = PS.PackageID
    JOIN v_PackageStatusDetailSumm AS PSD ON TSR.RefPackageID = PSD.PackageID
    JOIN v_Package as P on ps.PackageID = p.PackageID
    JOIN v_PackageStatusRootSummarizer n ON p.PackageID = n.PackageID
    WHERE TSR.PackageID = '$tasksequencePackageid' AND PS.PackageID NOT IN (SELECT
    PS.PackageID FROM v_PackageStatus AS PS 
    JOIN v_TaskSequencePackageReferences AS TSR ON PS.PackageID = TSR.RefPackageID 
    WHERE PS.PkgServer like '%$distributionPoint%') order by n.SourceSize"
    return executeSQLQuery
}

function distributePackages($packageid,$packagename,$packagetype)
{
    $status=$null
    $status=checkPackageStatus $packageid
    Writelog $status
    writeLog "Status of $packageid : $($status.installstatus)" 
    if(($null -eq $status) -or ($status.installstatus -ne 'Package Installation complete') -or ($status.installstatus -ne 'Content updating') -or ($status.installstatus -ne 'Retrying package installation'))
    {
        writelog "Starting distribution of $packagename." 
        try 
        {
            switch ($packagetype) 
            {
                0 { Start-CMContentDistribution -PackageId $packageid -DistributionPointName $distributionpoint }
                3 { Start-CMContentDistribution -DriverPackageId $packageid -DistributionPointName $distributionpoint }
                5 { Start-CMContentDistribution -DeploymentPackageId $packageid -DistributionPointName $distributionpoint }
                8 
                { 
                    $appid=$null
                    $appid=(Get-CMApplication -Fast -Name "*$packagename*").ci_id
                    if($appid -is [array])
                    {
                        $appid=$appid[0]
                    }
                    writelog "Application ID: $appid" 
                    Start-CMContentDistribution -ApplicationId "$appid" -DistributionPointName $distributionpoint 
                }
                257 { Start-CMContentDistribution -OperatingSystemImageId $packageid -DistributionPointName $distributionpoint}
                258 { Start-CMContentDistribution -BootImageId $packageid -DistributionPointName $distributionpoint}
                259 { Start-CMContentDistribution -OperatingSystemInstallerId $packageid -DistributionPointName $distributionpoint}
                Default 
                { 
                    $global:exitloop=$true
                    writelog "Package $packageid type is unknown. Hence skipped package distribution."  
                }
            }            
        }
        catch [System.Management.Automation.ItemNotFoundException] 
        {
            writelog "Package $packagename is not found in MECM." 
            $global:exitloop=$true
        }
        catch 
        {
            writelog "Error ocuured during package distribution $_." 
            $global:exitloop=$true
        }
    }
    else 
    {
        writelog "$packagename is already distributed." 
    }
}

Add-Type -AssemblyName System.speech
$global:speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$global:tasksequencePackageid=""
$global:distributionPoint=""
$global:SiteCode = ""
$global:ProviderMachineName = "" 
$global:dbserver=""
$global:dbcatalog=""
$global:query=""
$global:logfile="c:\temp\$distributionPoint-$tasksequencepackageid-distributionStatus.log"

$global:sqlcmd=$global:sqlconn=$null
$maxSize=100000
$isweekend=(get-date).DayOfWeek.value__ -in (6,0)
if($isweekend)
{
    $maxSize=500000
}
writelog "**************************************************************************************" 
writelog "Task Sequence PackageID : $tasksequencePackageid `t Distribution Point : $distributionPoint" 
if(($null -ne $tasksequencepackageid) -and ($null -ne $distributionPoint))
{
    intializeSQLConnection
    writelog "Checking missing packages of Task Sequence PackageID : $tasksequencePackageid on Distribution Point : $distributionPoint" 
    $qData=checkMissingPackages
    
    writelog "List of packages missing : $($qdata.name.count) `n $($qdata.name | ForEach-Object {$_;write-host ''})" 
    if($qdata.name.count -ne 0)
    {
        #first distribute packages less than equal to 100MB
       $quickDistribute=$qData | Where-Object { $_.sourcesize -lt $maxSize}  
       writelog "First 100 packages less than 100MB. `n $($quickDistribute.name)" 
        foreach($p in $quickDistribute)
        {
            distributePackages $p.PackageID $p.name $p.PackageType
        }
        Start-Sleep $(120*$($quickDistribute.count))
        $p=$null
        foreach($p in $qData)
        {
            $global:exitloop=$false
            $counter=1
            writelog "------------------------------------------------------------------------------------" 
            distributePackages $p.PackageID $p.name $p.PackageType
            if(!$exitloop)
            {
                do
                {
                    $sleeptime=$([int]($p.sourcesize/1000))
                    if($sleeptime -lt 300)
                    {
                        $sleeptime=300
                    }
                    else 
                    {
                        $sleeptime=$sleeptime*2
                    }
                    writelog "Package $($p.packageid) being distributed. Sleppeing for $sleeptime. Attempt : $counter" 
                    Start-Sleep $sleeptime
                    $status=$null
                    $status=checkPackageStatus $p.packageid
                    $counter++
                }
                while($status.installstatus -ne 'Package Installation complete' -and $counter -lt 6)
                writelog "Package $($p.packageid) distributed successfully." 
            }
            writelog "------------------------------------------------------------------------------------" 
        }
    }
    else 
    {
        writelog "No package to distribute to $distributionpoint." 
    }
    $global:sqlconn.close()
    set-location c:
}
else 
{
    writelog "Distribution point and task sequence package id information is missing." 
}
writelog "**************************************************************************************" 

