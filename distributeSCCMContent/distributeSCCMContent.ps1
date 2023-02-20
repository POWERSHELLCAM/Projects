function InitializeSCCM  
{  
    # Customizations  
    $initParams = @{}  
    
    # Import the ConfigurationManager.psd1 module   
    if($null -eq (Get-Module ConfigurationManager)) {  
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams   
    }  
    
    # Connect to the site's drive if it is not already present  
    if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) 
    {  
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams  
    }  
    
    # Set the current location to be the site code.  
    Set-Location "$($SiteCode):\" @initParams  
}  

function intializeSQLConnection
{
    $global:sqlConn = New-Object System.Data.SqlClient.SqlConnection
    $global:sqlConn.ConnectionString = "Server=$dbserver;Integrated Security=true;Initial Catalog=$dbcatalog"
    $global:sqlConn.open()
    $global:sqlcmd = New-Object System.Data.SqlClient.SqlCommand
    $global:sqlcmd.Connection = $sqlConn
    Write-Output "$(get-date -format G)> DBServer : $dbserver `t Database : $dbcatalog" | Out-File $logfile -Append
    InitializeSCCM
}

function executeSQLQuery
{
    $list=$null
    $global:sqlcmd.CommandText=$null
    $global:sqlcmd.CommandText = $query
    write-host $query
    $adp = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcmd
    $adp.SelectCommand.CommandTimeout=120
    $data = New-Object System.Data.DataSet
    $adp.Fill($data) | Out-Null
    $list=$data.tables
    $data=$adp=$null
    return $list
}

function checkPackageStatus($packageid)
{
    Write-Output "$(get-date -format G)> Checking package status of $packageid" | out-file $logfile -Append
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
    Write-Output "$(get-date -format G)> $status" | out-file $logfile -Append
    Write-Output "$(get-date -format G)> Status of $packageid : $($status.installstatus)" | out-file $logfile -Append
    if(($null -eq $status) -or ($status.installstatus -ne 'Package Installation complete') -or ($status.installstatus -ne 'Content updating') -or ($status.installstatus -ne 'Retrying package installation'))
    {
        Write-Output "$(get-date -format G)> Starting distribution of $packagename." | out-file $logfile -Append
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
                    Write-Output "$(get-date -format G)> Application ID: $appid" | out-file $logfile -Append
                    Start-CMContentDistribution -ApplicationId "$appid" -DistributionPointName $distributionpoint 
                }
                257 { Start-CMContentDistribution -OperatingSystemImageId $packageid -DistributionPointName $distributionpoint}
                258 { Start-CMContentDistribution -BootImageId $packageid -DistributionPointName $distributionpoint}
                259 { Start-CMContentDistribution -OperatingSystemInstallerId $packageid -DistributionPointName $distributionpoint}
                Default 
                { 
                    $global:exitloop=$true
                    Write-Output "$(get-date -format G)>Package $packageid type is unknown. Hence skipped package distribution." | out-file $logfile -Append 
                }
            }            
        }
        catch [System.Management.Automation.ItemNotFoundException] 
        {
            Write-Output "$(get-date -format G)> Package $packagename is not found in MECM." | out-file $logfile -Append
            $global:exitloop=$true
        }
        catch 
        {
            Write-Output "$(get-date -format G)> Error ocuured during package distribution $_." | out-file $logfile -Append
            $global:exitloop=$true
        }
    }
    else 
    {
        Write-Output "$(get-date -format G)> $packagename is already distributed." | out-file $logfile -Append
    }
}

$global:tasksequencePackageid="MP1003D4"
$global:distributionPoint="ngpsccmprddp1.corp.root.global"
$global:SiteCode = 'MP1'
$global:ProviderMachineName = 'mpgclwpsh0031.corp.root.global' 
$global:dbserver="MPGCLWPSH0030\GS_SCCMPRD"
$global:dbcatalog="cm_mp1"
$global:query=""
$global:logfile="c:\temp\$distributionPoint-$tasksequencepackageid-distributionStatus.log"
$global:sqlcmd=$global:sqlconn=$null
Write-Output "`n $(get-date -format G)> **************************************************************************************" | Out-File $logfile -Append
Write-Output "$(get-date -format G)> Task Sequence PackageID : $tasksequencePackageid `t Distribution Point : $distributionPoint" | Out-File $logfile -Append
if(($null -ne $tasksequencepackageid) -and ($null -ne $distributionPoint))
{
    intializeSQLConnection
    Write-Output "$(get-date -format G)> Checking missing package of Task Sequence PackageID : $tasksequencePackageid on Distribution Point : $distributionPoint" | Out-File $logfile -Append
    $qData=checkMissingPackages
    
    Write-Output "$(get-date -format G)> List of packages($($qdata.name.count)) missing : `n $($qdata.name | ForEach-Object {$_;write-host ''})" | Out-File $logfile -Append
    if($qdata.name.count -ne 0)
    {
        #first distribute packages less than equal to 100MB
       $first100=$qData | Where-Object { $_.sourcesize -lt 100000}  
       Write-Output "First 100 packages less than 100MB. `n $($first100.name)" | Out-File $logfile -Append
        foreach($p in $first100)
        {
            distributePackages $p.PackageID $p.name $p.PackageType
        }
        Start-Sleep $(120*$($first100.count))
        $p=$null
        $rest=$qData #| Where-Object { $_.sourcesize -ge 100000} 
        foreach($p in $rest)
        {
            $global:exitloop=$false
            $counter=1
            Write-Output "------------------------------------------------------------------------------------" | Out-File $logfile -Append
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
                    Write-Output "$(get-date -format G)> Package $($p.packageid) being distributed. Sleppeing for $sleeptime. Attempt : $counter" | Out-File $logfile -Append
                    Start-Sleep $sleeptime
                    $status=$null
                    $status=checkPackageStatus $p.packageid
                    $counter++
                }
                while($status.installstatus -ne 'Package Installation complete' -and $counter -lt 6)
                Write-Output "$(get-date -format G)> Package $($p.packageid) distributed successfully." | Out-File $logfile -Append
            }
            Write-Output "------------------------------------------------------------------------------------" | Out-File $logfile -Append
        }
    }
    else 
    {
        Write-Output "$(get-date -format G)> No package to distribute to $distributionpoint." | Out-File $logfile -Append
    }
    $global:sqlconn.close()
    set-location c:
}
else 
{
    Write-Output "$(get-date -format G)> Distribution point and task sequence package id information is missing." | Out-File $logfile -Append
}
Write-Output "`n $(get-date -format G)> **************************************************************************************" | Out-File $logfile -Append

