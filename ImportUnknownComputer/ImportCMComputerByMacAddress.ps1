
$init=
{
    
    function writemsg($m)
    {
        Write-Output "$m ($(get-date -format G))" | Out-File $importlog -Append
    }

    function checkDeviceMembership
    {
        $attempt=0
        $deviceAddedToCollection=$null
        writemsg "Checking device membership in image deployment group."
        do 
        {
            $attempt++
            $deviceAddedToCollection=$null
            if(isPartOfImageCollection)
            {
                $deviceAddedToCollection=$true
                $global:addedtoimagecollection=$true
                break
            }
            else
            {
                writemsg "Sleeping for 120 seconds which checking image collection membership."
                Start-Sleep 120
            }
        } while ($true -and $attempt -le 20)
        if(!$deviceAddedToCollection)
        {
            writemsg "Failed to check the membership. Please contact admin to ensure the device is part of collection and then proceed with deployment."
            return $false
        }
        else 
        {
            writemsg "Successfully added the device to imaging group."
            return $true
        }
    }
    function isPartOfRootCollection
    {
        if(Get-WmiObject -ComputerName $SCCMServerName -Namespace $SCCMNameSpace -Query "select * from SMS_FullCollectionMembership where CollectionID = 'SMS00001' and resourceid ='$resourceid'")
        {
            return $true
        }
        else 
        {
            return $false
        }
    }
    function isPartOfImageCollection
    {
        if(Get-WmiObject -ComputerName $SCCMServerName -Namespace $SCCMNameSpace -Query "select * from SMS_FullCollectionMembership where CollectionID = '$collectionid' and resourceid ='$resourceid'")
        {
            return $true
        }
        else 
        {
            return $false
        }
    }
    function addDeviceToSCCMCollection
    {
        try 
        {
            $collection=Get-WMIObject -Computer $SCCMServerName -Namespace $SCCMNameSpace -class SMS_Collection -filter "CollectionID='$collectionId'"
            $addToCollectionParameters = $collection.GetmethodParameters("AddMembershipRule")
            $rule = ([WMIClass]("\\$SCCMServerName\$SCCMNameSpace" + ":SMS_CollectionRuleDirect")).CreateInstance()
            $rule.ResourceClassName = "SMS_R_System"
            $rule.ResourceID = $resourceId   
            $addToCollectionParameters.CollectionRule = $rule
            $collection.InvokeMethod("AddMembershipRule", $addToCollectionParameters, $null)
            writemsg "Successfully initiated device $LocalComputername membership in image deployment group. $?" 1
            checkDeviceMembership
        }
        catch 
        {
            writemsg "Device $LocalComputername having resource id $resourceid failed to add in image deployment group. $?"
            writemsg $_
            return $false
            $global:mailstatus="Failed"
        }
    }
    function importDeviceToSCCM ($LocalComputerName,$id)
    {
        try 
        {
            $WMINameSpace = Get-WMIObject -List -ComputerName $SCCMServerName -NameSpace $SCCMNameSpace -class "SMS_Site"
            writemsg "Creating a new record of a device $LocalComputerName with Unique ID $id."
            $NewComputerAccountDetail = $WMINameSpace.psbase.GetMethodParameters("ImportMachineEntry")
            if($id.length -eq 17)
            {
                $NewComputerAccountDetail.MACAddress = $id
            }
            else 
            {
                $NewComputerAccountDetail.SMBIOSGUID = $id
            }
            $NewComputerAccountDetail.NetbiosName = $LocalComputerName
            $NewComputerAccountDetail.OverwriteExistingRecord = $True
            $resource = $WMINameSpace.psbase.InvokeMethod("ImportMachineEntry",$NewComputerAccountDetail,$null)
            $global:resourceid=$($resource.resourceid)
            writemsg "Device $LocalComputername record created successfully."
            writemsg "Resource ID = $resourceid."
            writemsg "Please wait for some time to get the record refreshed in database."
            while(!(isPartOfRootCollection))
            {
                writemsg "Sleeping for 360 seconds while checking root collection membership."
                Start-Sleep 360
            }
            addDeviceToSCCMCollection
        }
        catch 
        {
            writemsg $_
            $global:mailstatus="Failed"
        }
    }
    function requestImport($techemail,$uniqueid,$hostname,$collectionid,$user,$csvfilepath)
    {
        $global:mailstatus=""
        $global:mailmessage=""
        $global:resourceid=$null
        $LocalComputerName=$hostname
        $username=$user
        $extramsg=$null
        $global:addedtoimagecollection=$false
        $SCCMServerName=''
        $SCCMNameSpace='root\sms\site_'
        $sharedpath="c:\temp\ImportUnknownComputer"
        $importlog="$sharedpath\logs\$hostname.log"
        writemsg "CollectionId: $collectionid"
        writemsg "User Name : $username"
        $multiplerecords=$false
        $comment=""
        $collectionname=(Get-WMIObject -Computer $SCCMServerName -Namespace $SCCMNameSpace -class SMS_Collection -filter "CollectionID='$collectionId'").name
        if($uniqueid.length -eq 17)
        {
            $existingdevice=Get-WmiObject -class "SMS_R_System" -namespace $SCCMNameSpace -Filter "MACAddresses = '$($uniqueid)'" -ComputerName $SCCMServerName
        }
        else 
        {
            $existingdevice=Get-WmiObject -class "SMS_R_System" -namespace $SCCMNameSpace -Filter "SMBIOSGUID = '$($uniqueid)'" -ComputerName $SCCMServerName

        }

        if($existingdevice -is [array])
        {
            $multiplerecords=$true
        }
        
        if(!$existingdevice)
        {
            try 
            {
                importDeviceToSCCM $hostname $uniqueid
                if($resourceid)
                {
                    if($addedtoimagecollection)
                    {
                        $global:mailstatus="Completed"
                        $global:mailmessage="<p style='color:green'>Registered. Proceed with imaging.</p>"
                    }
                    else 
                    {
                        $global:mailstatus="Failed"
                        $global:mailmessage="<p style='color:red'>Failed to register. Connect GT MECM Platform Team for issue and solution.</p>"
                    }
                }
            }
            catch 
            {
                writemsg $_
                $global:mailstatus="Failed"
                $global:mailmessage="<p style='color:red'>Failed to register. Connect GT MECM Platform Team for issue and solution.</p>"
                $comment="Failed to create record in MECM."
            } 
        }
        else 
        {
            $existingname=$($existingdevice.name)
            if($existingname -ne $hostname)
            {
                $comment="Existing device name : $existingname"
            }
            writemsg "A device with name $existingname already exist with UniqueID $uniqueid"
            if($multiplerecords)
            {
                $global:resourceid=$($existingdevice[0].resourceid)
                $extramsg="**Duplicate records found with UniqueID $uniqueid.**"
                writemsg $extramsg
                $comment=$extramsg
            }
            else 
            {
                $global:resourceid=$($existingdevice.resourceid)
            }
            
            if(isPartOfImageCollection)
            {
                $global:mailstatus="Completed"
                writemsg "A device with name $hostname already part of imaging collection $collectionid."
                $global:mailmessage="<p style='color:green'>Already Registered. Proceed with imaging.</p>"
            }
            else 
            {
                addDeviceToSCCMCollection
                if($addedtoimagecollection)
                {
                    $global:mailstatus="Completed"
                    $global:mailmessage="<p style='color:green'>Already Registered. Proceed with imaging.</p>"
                }
                else 
                {
                    $global:mailstatus="Failed"
                    $global:mailmessage="<p style='color:red'>Failed to register. Connect GT MECM Platform Team for issue and solution.</p>"
                    $comment="Failed to add in imaging collection."
                }
            }
        }
        if($mailmessage -eq "")
        {
            $global:mailmessage="<p style='color:red'>Failed to register .Connect GT MECM Platform Team for issue and solution.</p>"
        }
        #sendmail $mailmessage $techemail
        writemsg "------------------------------------------------------------------------------------------------------------------------------------"
        if($mailstatus -ne "completed")
        {
            #return $csvfilepath
        }
        $myReport=$null
        $myReport = [pscustomobject]@{
            ComputerName = $hostname
            Username=$username
            ResourceID = $resourceid
            UniqueID = $uniqueid
            CollectionID=$collectionid
            CollectionName=$collectionname
            Status    = $mailstatus
            Email= $techemail
            Message=$mailmessage
            Comment=$comment
            CSVFile=$csvfilepath
        }
        return $myReport
    }
}

function sendmail($techemail,$username,$table)
    {
        $TextInfo = (Get-Culture).TextInfo
        $Mailbox = ""
        $SMTP = ""  
        $Subject = "Your device import status in MECM "
        $body = "<HTML><HEAD></HEAD>"
        $body+='<BODY style="font-size:11pt;font-family:Calibri;color:black">'
        $body += "<br>"
        $body += "Hello $($TextInfo.ToTitleCase($username)),<br>"
        $body += "<br>"
        $body += "Thanks for your request. Please find the import status of device/s requested."
        $body += "<br>"
        $body += "$table"
        $body += "<br>"
        $body += "Regards,<br>GT MECM Platform Team"
        $body += "</body>"
        
        #region send Windows Deployment Email with additional information
        $params = $null
        $params = @{ 
            To = $techemail
            Subject = $Subject
            Body = $body 
            SmtpServer = $SMTP
            From = $Mailbox
            BodyAsHtml = $true 
        } 
        Send-MailMessage @params -ErrorAction SilentlyContinue
        writemsg "Mail send status : $?"
        #endregion send Windows Deployment Email with additional information

    }

function writemsg($m)
{
    Write-Output "$m ($(get-date -format G))" | Out-File $importlog -Append
}
#--CSS formatting
$test=@'
<style type="text/css">
 h1, h5,h2, th { text-align: left; font-family: Segoe UI;font-size: 13px;}
table { margin: left; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; font-size: 12px;}
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
'@
$global:mailstatus=$null
$global:mailmessage=""
$global:resourceid=$null
$global:addedtoimagecollection=$false
$sharedpath="c:\temp\ImportUnknownComputer"
$importlog="$sharedpath\ImportCMComputerByMacAddress.log"
if((Get-Item "$importlog").length/1MB -gt 5)
{
    Move-Item $importlog -Destination "$sharedpath\_backup\_log\ImportCMComputerByMacAddress_$(Get-Date -Format 'yyyyMMdd-HHMMss').log"
}
Get-ChildItem -Path "$sharedpath\logs\*.log" | Where-Object {($_.LastWriteTime -lt (Get-Date).AddDays(-30))} | Remove-Item -Force -ErrorAction SilentlyContinue
$csvdumps="$sharedpath\PSCode\csvdumps"
$collectionlist=Import-Csv -Path "$sharedpath\db\coutryCollList.csv"
$csvlist=Get-ChildItem -Path $csvdumps -Filter *.csv #| Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays(-2) } | Sort-Object LastWriteTime -Descending
$finalreport=@()


if($null -ne $csvlist)
{
    writemsg "------------------------------------------------------------------------------------------------------------------------------------"
    writemsg "$csvlist"
    foreach($list in $csvlist)
    {
        writemsg "Processig $list"
        $csvimportedList=$existingdevice=$null
        $csvimportedList=Import-Csv -Path $($list.fullname)
        if($csvimportedList)
        {
            foreach($csvimported in $csvimportedList)
            {
                $collectionid=$global:userid=$null
                $global:mailstatus=""
                $collectionid=($collectionlist -match $($csvimported.country)).collectionid
                $username=$($csvimported.email).split('.')[0]
                
                $requested=$false
                if($($csvimported.guid))
                {
                    writemsg "A thread initiated for processing $($csvimported.guid)"
                    Start-Job -ScriptBlock {requestImport $args[0] $args[1] $args[2] $args[3] $args[4] $args[5]} -InitializationScript $init -ArgumentList ($($csvimported.email), $($csvimported.guid), $($csvimported.computer),$collectionid,$username,$($list.fullname))
                    $requested=$true
                }
                if($($csvimported.MACAddress) -and !$requested)
                {
                    writemsg "A thread initiated for processing $($csvimported.MACAddress)"
                    Start-Job -ScriptBlock {requestImport $args[0] $args[1] $args[2] $args[3] $args[4] $args[5]} -InitializationScript $init -ArgumentList ($($csvimported.email), $($csvimported.MACAddress), $($csvimported.computer),$collectionid,$username,$($list.fullname))
                }
                
            }
        }
        else 
        {
            writemsg "Import CSV dont have any entry."
        }  
    }
    Get-Job | Wait-Job  
    writemsg "All thread finished processing."
    $jobs=get-job | Receive-Job
    $jobs = $jobs  | Group-Object -Property email
    $jobs
    foreach($j in $jobs)
    {
        $mailto=$($j.name)
        $name=$($j.name).split('.')[0]
        $table=$($j.group) | Select-Object computername,resourceid,uniqueid,collectionid,collectionname,status,message,comment | ConvertTo-Html -Head $test
        $table=[Net.WebUtility]::Htmldecode($table)
        $table
        sendmail $mailto $name $table

    }
    writemsg $jobs
     Get-Job | remove-Job  
    $filelist = $jobs | Select-Object -Unique
    writemsg "List of unique .CSV files entry."
    $filelist | ForEach-Object { writemsg $_}
    foreach($list in $csvlist)
    {
        if($filelist -notcontains $($list.fullname))
        {
            writemsg "Moving $($list.fullname) to '$sharedpath\_backup\_csv'"
            Move-Item $($list.fullname) -Destination "$sharedpath\_backup\_csv" -force
            writemsg "------------------------------------------------------------------------------------------------------------------------------------"
        }
    }
}
else 
{
    write-host "No item to process. Exiting..."
}



