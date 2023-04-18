function displayMsg($m)
{
    $global:msg.content=""
    if($m -match 'finished')
    {
        $global:msg.foreground='green'
    }
    if($m -match 'started')
    {
        $global:msg.foreground='black'
    }
    $global:msg.content=$m
}

function fillgridview($data)
{
    $($XMLForm.FindName('ResultDataGrid')).ItemsSource=$data
}

function updateHTML
{
    param ($strPath)
    IF(Test-Path $strPath)
    { 
        Remove-Item $strPath
    }   
}

function generateHTMLReport
{
    $applist=$filelist=$folderlist=$reglist=$customlist=$null
    $applist=$filelist=$folderlist=$reglist=$customlist=@()
    $checklistxml.Check.Applications.App | ForEach-Object{
        $packname=$packexist=""
        $packname=$null
        $packname=$(Get-Package "*$_*")
        if($null -eq $packname){$packexist="Application not installed"}else{$packexist="Application installed"}
        if($null -ne $packname)
        {
            foreach($app in $packname) 
            {
                $applist+=[PSCustomObject]@{
                    Name = $_
                    'Full Name'=$($app.name)
                    'Version'=$($app.Version)
                    'Status'= $packexist
                }
            }
        }
        else 
        {
            $applist+=[PSCustomObject]@{
                Name = $_
                'Full Name'="NA"
                'Version'="NA"
                'Status'= $packexist
            } <# Action when all if and elseif conditions are false #>
        }
    }
    $checklistxml.Check.Settings.File.Path | ForEach-Object {
        $filelist+=[PSCustomObject]@{
            Path = $_
            'Status'=if(Test-Path $_ -ErrorAction SilentlyContinue){"File exists."}else{"File does not exist."}
        }
    }
    $checklistxml.Check.Settings.Folder.Path | ForEach-Object {
        $folderlist+=[PSCustomObject]@{
            Path = $_
            'Status'=if(Test-Path $_ -ErrorAction SilentlyContinue){"Folder exists."}else{"Folder does not exist."}
        }
    }
    
    $checklistxml.Check.Settings.Registery.Reg | ForEach-Object {
        $reglist+=[PSCustomObject]@{
            'Reistry Path' = $_.path
            'Registry Key' = $_.key
            'Registry Value' = $_.val
            'Status' = if((Invoke-Expression $("`$(Get-ItemProperty $($_.path))."+"$($_.key)")) -eq $_.val){"Reg key and value matches."}else{"Reg key and value does not matche."}
        }
    }  
    
    $checklistxml.Check.Settings.Custom.set | ForEach-Object {
        $customlist+=[PSCustomObject]@{
            'Setting Name'=$_.name
            'Expected Output'=$_.Output
            'Real Output' = Invoke-Expression $($_.command)
            'Status' = if($(Invoke-Expression $($_.command)) -eq $_.Output){'Output is matching.'}else{'Output is not matching.'}
        }
    }

    $applist | ConvertTo-html  -Head $test -Body "<h2>Applications Validation</h2>" >> "$strPath"
    $filelist | ConvertTo-html  -Head $test -Body "<h2>Files Validation</h2>" >> "$strPath"
    $folderlist | ConvertTo-html  -Head $test -Body "<h2>Folders Validation</h2>" >> "$strPath"
    $reglist | ConvertTo-html  -Head $test -Body "<h2>Registry Validation</h2>" >> "$strPath"
    $customlist | ConvertTo-html  -Head $test -Body "<h2>Custom Validation</h2>" >> "$strPath"
    #Launching HTML generated report
    Invoke-Item $strPath
}


#Load Assembly and Library
Add-Type -AssemblyName PresentationFramework

#XAML form designed using Vistual Studio
$sharedpath=$PSScriptRoot
$bannerfile = "$sharedpath\_logo\banner.png"
[xml]$Form = Get-Content "$sharedpath\_xml\WPF GUI.xml"
$checklistxml= [xml](Get-Content "$sharedpath\_xml\checklistXML.xml")
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

$applist=$filelist=$folderlist=$reglist=$customlist=$null
$applist=$filelist=$folderlist=$reglist=$customlist=@()

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

$ReportTitle="Autopilot Post checklist"
$strPath = "$sharedpath\_reports\$ReportTitle.html" 
updateHTML $strPath

ConvertTo-Html -Head $test -Title $ReportTitle -Body "<h1> Computer  $($env:COMPUTERNAME) validation report</h1>" >  "$strPath"

#Create a form
$XMLReader = (New-Object System.Xml.XmlNodeReader $Form)
$XMLForm = [Windows.Markup.XamlReader]::Load($XMLReader)

#Load Controls
$XMLForm.FindName('logo').Source = $bannerfile
$Progress = $XMLForm.FindName('progress')
$ResultDataGrid = $XMLForm.FindName('ResultDataGrid')
$clock = $XMLForm.FindName('clock')
$diskdisplay = $XMLForm.FindName('disk')
$global:msg = $XMLForm.FindName('msg')
$disk=Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object @{'Name' = 'FreeSpace'; Expression= { [int]($_.FreeSpace / 1GB) }},@{'Name' = 'TotalSize'; Expression= { [int]($_.size / 1GB) }}
$totalspace=$totalfree=0
foreach($d in $disk)
{
    $totalspace+=$d.TotalSize
    $totalfree+=$d.FreeSpace
}
$diskutilization=[int]((($totalspace-$totalfree)*100)/$totalspace)
$diskdisplay.content="Disk Utilization: $diskutilization%"
$Progress.value=$diskutilization
$XMLForm.FindName('updates').add_click({
    fillgridview $(get-wmiobject -class win32_quickfixengineering | Select-Object HotFixID,Description,InstalledOn,Installedby)
})

$XMLForm.FindName('regional').add_click({
    $lname=$null
    (Get-WinUserLanguageList).Localizedname | ForEach-Object {$lname=$_+","+$lname}
    fillgridview @($([PSCustomObject]@{
        'Language Packs' = $lname.Substring(0,$lname.Length-1)
        'Display Language'=(GET-WinSystemLocale).displayname
        'Time Zone'=$((get-timezone).standardname)+":"+$((get-timezone).displayname)
    }))
})

$XMLForm.FindName('checkapp').add_click({
    displayMsg "Started validating applications. Please wait...."
    $checklistxml.Check.Applications.App | ForEach-Object{
        $packname=$packexist=""
        $packname=$null
        $packname=$(Get-Package "*$_*")
        if($null -eq $packname){$packexist="Application not installed"}else{$packexist="Application installed"}
        if($null -ne $packname)
        {
            foreach($app in $packname) 
            {
                $applist+=[PSCustomObject]@{
                    Name = $_
                    'Full Name'=$($app.name)
                    'Version'=$($app.Version)
                    'Status'= $packexist
                }
            }
        }
        else 
        {
            $applist+=[PSCustomObject]@{
                Name = $_
                'Full Name'="NA"
                'Version'="NA"
                'Status'= $packexist
            } <# Action when all if and elseif conditions are false #>
        }
    }
    fillgridview $applist
    displayMsg "Finished validating applications."
})

$XMLForm.FindName('checkfile').add_click({
    displayMsg "Started validating files. Please wait...."
    $checklistxml.Check.Settings.File.Path | ForEach-Object {
        $filelist+=[PSCustomObject]@{
            Path = $_
            'Status'=if(Test-Path $_ -ErrorAction SilentlyContinue){"File exists."}else{"File does not exist."}
        }
    }
    fillgridview $filelist
    displayMsg "Finished validating files."
})

$XMLForm.FindName('checkfolder').add_click({
    displayMsg "Started validating folders. Please wait...."
    $checklistxml.Check.Settings.Folder.Path | ForEach-Object {
        $folderlist+=[PSCustomObject]@{
            Path = $_
            'Status'=if(Test-Path $_ -ErrorAction SilentlyContinue){"Folder exists."}else{"Folderdoes not exist."}
        }
    }
    fillgridview $folderlist
    displayMsg "Finished validating folders."
})

$XMLForm.FindName('checkreg').add_click({
    displayMsg "Started validating Registry. Please wait...."
    $checklistxml.Check.Settings.Registery.Reg | ForEach-Object {
        $reglist+=[PSCustomObject]@{
            'Reistry Path' = $_.path
            'Registry Key' = $_.key
            'Registry Value' = $_.val
            'Status' = if((Invoke-Expression $("`$(Get-ItemProperty $($_.path))."+"$($_.key)")) -eq $_.val){"Reg key and value matches."}else{"Reg key and value does not matche."}
        }
    } 
    fillgridview $reglist
    displayMsg "Finished validating registry."
})

$XMLForm.FindName('checkcustom').add_click({
    displayMsg "Started validating custom settings. Please wait...."
    $checklistxml.Check.Settings.Custom.set | ForEach-Object {
        $customlist+=[PSCustomObject]@{
            'Setting Name'=$_.name
            'Expected Output'=$_.Output
            'Real Output' = Invoke-Expression $($_.command)
            'Status' = if($(Invoke-Expression $($_.command)) -eq $_.Output){'Output is matching.'}else{'Output is not matching.'}
        }
    }
    fillgridview $customlist
    displayMsg "Finished validating custom settings."
})

$XMLForm.FindName('activation').add_click({
    displayMsg "Processing activation information. Please wait...."
    fillgridview @($(Get-WmiObject -query 'select * from SoftwareLicensingService' | Select-Object OA3xOriginalProductKey,OA3xOriginalProductKeyDescription,KeyManagementServiceDnsPublishing,KeyManagementServiceListeningPort,KeyManagementServicePort,RemainingWindowsReArmCount))
    displayMsg "Activation information."
})

$XMLForm.FindName('domain').add_click({
    fillgridview @($(Get-WmiObject -Class Win32_ComputerSystem | Select-Object PrimaryOwnerName,Partofdomain, Domain, HypervisorPresent))
    displayMsg "Domain information."
})

$XMLForm.FindName('report').add_click({generateHTMLReport})
$XMLForm.FindName('applications').add_click({
    displayMsg "Started processing installed applications. Please wait...."
    $array=@()
    $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
    $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$env:COMPUTERNAME) 
    $regkey=$reg.OpenSubKey($UninstallKey) 
    $subkeys=$regkey.GetSubKeyNames() 
        foreach($key in $subkeys)
        {
            $key
            $thisKey=$UninstallKey+"\\"+$key 
            $thisSubKey=$reg.OpenSubKey($thisKey)
            if($null -ne $($thisSubKey.GetValue("DisplayName")))
            {
                $obj = [PSCustomObject]@{
                    'Name' = $($thisSubKey.GetValue("DisplayName"))
                    'Version' = $($thisSubKey.GetValue("DisplayVersion"))
                    'Publisher' = $($thisSubKey.GetValue("Publisher"))
                    'Install Date'= if($($thisSubKey.GetValue("InstallDate"))){[datetime]::Parseexact($($thisSubKey.GetValue("InstallDate")),"yyyyMMdd", $null)}
                }
                $array += $obj
            }
        }
    fillgridview $array
    displayMsg "Installed Applications."
})

$XMLForm.FindName('drivers').add_click({
    fillgridview $(Get-WmiObject Win32_PNPEntity | Select-Object Name,Present,Status,ConfigManagerErrorCode)
    devmgmt.msc
    displayMsg "Driver status."
})

$XMLForm.FindName('bitlocker').add_click({
    fillgridview $(Get-BitLockerVolume | Select-Object VolumeType,MountPoint, VolumeStatus, EncryptionPercentage,CapacityGB,ProtectionStatus)
    displayMsg "Bitlocker Information"
})

$($XMLForm.FindName('sysgrid')).ItemsSource=@([pscustomobject]@{
    Name=$env:COMPUTERNAME
    OS=(Get-WMIObject win32_operatingsystem).Caption
    'OS Version'=(Get-WMIObject win32_operatingsystem).Version
    Make=(Get-WMIObject -class Win32_ComputerSystem).Manufacturer
    Model=(Get-WMIObject -class Win32_ComputerSystem).model
    Serial=(Get-WmiObject win32_bios).serialnumber
    'BIOS Version'=(Get-WmiObject win32_bios).Version
})

$timer1 = New-Object 'System.Windows.Forms.Timer'
#$timer1 = $XMLForm.FindName('timer')
$timer1_Tick={
    $clock.Content= (Get-Date).ToString("dd-MM-yyyy HH:mm:ss")
}
$timer1.Enabled = $True
$timer1.Interval = 1
$timer1.add_Tick($timer1_Tick)
#Show XMLform
$XMLForm.ShowDialog()
