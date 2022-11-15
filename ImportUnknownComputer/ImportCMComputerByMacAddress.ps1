
$init=
{
    function sendmail($message,$techemail)
    {
        $TextInfo = (Get-Culture).TextInfo
        $Mailbox = ""
        $SMTP = ""  
        $Subject = "Device $LocalComputername imported in MECM "
        $body = "<HTML><HEAD></HEAD>"
        $body+='<BODY style="font-size:11pt;font-family:Calibri;color:black">'
        $body += "<br>"
        $body += "Hello $($TextInfo.ToTitleCase($username)),<br>"
        $body += "<br>"
        $body += $message
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
            $global:mailstatus="failed"
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
            $global:mailstatus="failed"
        }
    }
    function requestImport($techemail,$uniqueid,$hostname,$collectionid,$user,$csvfilepath)
    {
        $global:mailstatus=""
        $global:mailmessage=""
        $global:resourceid=$null
        $LocalComputerName=$hostname
        $username=$user
        $global:addedtoimagecollection=$false
        $SCCMServerName='<sccm primary>'
        $SCCMNameSpace='root\sms\site_MP1'
        $sharedpath="c:\temp\ImportUnknownComputer"
        $importlog="$sharedpath\logs\$hostname.log"
        writemsg "CollectionId: $collectionid"
        writemsg "User Name : $username"
        $multiplerecords=$false
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
                        $global:mailstatus="completed"
                        $global:mailmessage="Your request to add device <b><i>$hostname</i></b> having Unique ID <b><i>$uniqueid</i></b> is processed successfully. You may now proceed with the image deployment."
                    }
                    else 
                    {
                        $global:mailstatus="Failed"
                        $global:mailmessage="Your request to add device <b><i>$hostname</i></b> having Unique ID <b><i>$uniqueid</i></b> is failed. Please connect SOE Team for solution."
                    }
                }
            }
            catch 
            {
                writemsg $_
                $global:mailstatus="failed"
                $global:mailmessage="Your request to add device <b><i>$hostname</i></b> having Unique ID <b><i>$uniqueid</i></b> is failed. Please connect SOE Team for solution."
            } 
        }
        else 
        {
            writemsg "A device with name $hostname already exist with UniqueID $uniqueid"
            $extramsg=$null
            if($multiplerecords)
            {
                $global:resourceid=$($existingdevice[0].resourceid)
                $extramsg="**Duplicate records found with UniqueID $uniqueid. Please forward this mail to SOE engineer.**"
                writemsg $extramsg
            }
            else 
            {
                $global:resourceid=$($existingdevice.resourceid)
            }
            
            if(isPartOfImageCollection)
            {
                $global:mailstatus="completed"
                writemsg "A device with name $hostname already part of imaging collection $collectionid."
                $global:mailmessage="Your requested device <b><i>$hostname</i></b> having Unique ID <b><i>$uniqueid</i></b> is already registered with MECM. You may now proceed with the image deployment.<br><p style='color:red'>$extramsg</p>"
            }
            else 
            {
                addDeviceToSCCMCollection
                if($addedtoimagecollection)
                {
                    $global:mailstatus="completed"
                    $global:mailmessage="Your requested device <b><i>$hostname</i></b> having Unique ID <b><i>$uniqueid</i></b> is already registered with MECM. The requates is processed successfully. You may now proceed with the image deployment.<br><p style='color:red'>$extramsg</p>"
                }
                else 
                {
                    $global:mailstatus="failed"
                    $global:mailmessage="Your request to add device <b><i>$hostname</i></b> having Unique ID <b><i>$uniqueid</i></b> is failed. Please connect SOE Team for solution.<br><p style='color:red'>$extramsg</p>"
                }
            }
        }
        if($mailmessage -eq "")
        {
            $global:mailmessage="Your request to add device <b><i>$hostname</i></b> having Unique ID <b><i>$uniqueid</i></b> is failed. Please connect SOE Team for solution.<br><p style='color:red'>$extramsg</p>"
        }
        sendmail $mailmessage $techemail
        writemsg "------------------------------------------------------------------------------------------------------------------------------------"
        if($mailstatus -ne "completed")
        {
            return $csvfilepath
        }
    }
}
function writemsg($m)
{
    Write-Output "$m ($(get-date -format G))" | Out-File $importlog -Append
}
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
                writemsg "A thread initiated for processing $($csvimported.MACAddress)"
                if($($csvimported.MACAddress))
                {
                    Start-Job -ScriptBlock {requestImport $args[0] $args[1] $args[2] $args[3] $args[4] $args[5]} -InitializationScript $init -ArgumentList ($($csvimported.email), $($csvimported.MACAddress), $($csvimported.computer),$collectionid,$username,$($list.fullname))
                }
                if($($csvimported.guid))
                {
                    Start-Job -ScriptBlock {requestImport $args[0] $args[1] $args[2] $args[3] $args[4] $args[5]} -InitializationScript $init -ArgumentList ($($csvimported.email), $($csvimported.guid), $($csvimported.computer),$collectionid,$username,$($list.fullname))
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
    writemsg $jobs
    $filelist = $jobs | Select-Object -Unique
    writemsg "List of unique .CSV files entry."
    $filelist | ForEach-Object { writemsg $_}
    foreach($list in $csvlist)
    {
        if($filelist -notcontains $($list.fullname))
        {
            writemsg "Moving $($list.fullname) to '$sharedpath\_backup\_csv'"
            Move-Item $($list.fullname) -Destination "$sharedpath\_backup\_csv"
            writemsg "------------------------------------------------------------------------------------------------------------------------------------"
        }
    }
}
else 
{
    write-host "No item to process. Exiting..."
}



