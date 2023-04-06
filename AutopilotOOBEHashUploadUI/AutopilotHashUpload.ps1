
$init=
{
    function sethardwarehash($tag,$tenentid)
    {
        try 
        {
            Find-Script -Name Get-WindowsAutoPilotInfo
            Install-Script -Name Get-WindowsAutoPilotInfo -Force
            Set-ExecutionPolicy -ExecutionPolicy bypass -Force
            Get-WindowsAutoPilotInfo.ps1 -Online -GroupTag $tag -tenantid $tenentid
        }
        catch 
        {
            write-host "Error ocuured."
        }
    }
}

function addobject($objectType,$objectName,$x,$y,$objectText,$ObjectTabIndex,$x1,$y1) 
{   
    $object = New-Object System.Windows.Forms.$objectType 
	$object.Size = New-Object System.Drawing.size($x,$y)
    $object.Name = $objectName
	$object.Text = $objectText  
    $object.DataBindings.DefaultDataSourceUpdateMode = 0 
    if($objectType -ne "Form") 
    { 
            $object.TabIndex = $ObjectTabIndex 
            $object.Location = New-Object System.Drawing.Point($x1,$y1) 
    } 
    else
	{
		if($null -ne $x1)
		{
			$object.Location = New-Object System.Drawing.Point($x1,$y1) 
		}
		else 
		{
			$Object.StartPosition = 'CenterScreen'
		}
		$Object.BackColor="white"
	} 
	return $object
} 

function Append-ColoredLine {
    param( 
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Windows.Forms.RichTextBox]$box,
        [Parameter(Mandatory = $true, Position = 1)]
        [System.Drawing.Color]$color,
        [Parameter(Mandatory = $true, Position = 2)]
        [string]$text
    )
    $box.SelectionStart = $box.TextLength
    $box.SelectionLength = 0
    $box.SelectionColor = $color
    $box.AppendText($text)
    $box.AppendText([Environment]::NewLine)
}

function writemsg
{
    $global:index=$global:index+1
	Append-ColoredLine $msg $color "[$index] $message `r`n"
    $global:message=""
	$msg.Focus()
	$form.Refresh()
}

function readmsg($m,$c)
{
    if($c -eq 1)
    {
        $global:color="green"
    }
    if($c -eq 0)
    {
        $global:color="red"
    }
    if($c -eq 2)
    {
        $global:color="magenta"
    }
    $global:message="$m ($(get-date -format G))"
    writemsg
}



function findtag
{
    $tbgrouptag.text= $($importedcsv | Where-Object { $_.country -eq $($lbcountry.text)}).tag    
    readmsg "Autopilot Group Tag : $($tbgrouptag.text)." 2
    readmsg "Click 'Submit Button' to upload hardware hash and assign to group tag. Once upload completed, please close the Form and command prompt." 1
}

#variable declartion for windows Form .net object
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles();
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#Variable declaration for objects on Windows form like lable, textbox, button etc.
$lablesize_x=100
$lablesize_y=20
$labelfirstlocation_x=10
$labelfirstlocation_y=20
$boxfirstlocation_x =120
$boxfirstlocation_y=20
$boxsize_x=210
$boxsize_y=20
$formgap=30
$lablevariablelist=@('lregion','lcountry','lgrouptag')
$boxvariablelist=@('lbregion','lbcountry','tbgrouptag')
$labletextlist=@('Select Region:','Selecty Country:','Group Tag:')

$formtitle="<orgname> Autopilot Console v1.0"
$sharedpath=$PSScriptRoot
$tenentid=""
$img = [System.Drawing.Image]::Fromfile($(get-item "$sharedpath\_logo\banner.png"));
$importedcsv=Import-Csv "$sharedpath\_csv\grouptaglist.csv"
$global:message=""
$global:index=0

#region Form objects creation and customizations
#Main form design
$Form=addobject "Form" "form" 500 430 $formtitle $null $null $null
$Form.ControlBox = $false
$Form.TopMost = $true

#Groupbox
$groupBoxSingle = addobject "GroupBox" "groupbox" 400 150 "Online Hardware Hash Upload" $null 40 70
$groupBoxMsg = addobject "GroupBox" "groupbox" 400 150 "" $null 40 220

#logo
$pictureBox = addobject "PictureBox" "picture" 500 50 "ManpowerGroup Global SOE" $null 120 10
$pictureBox.Image = $img

for($i=0;$i -lt $lablevariablelist.Count;$i++)
{
    $indices=$i
    New-Variable -Name $lablevariablelist[$i] -Value $(addobject "Label" $lablevariablelist[$i] $lablesize_x $lablesize_y $labletextlist[$i] $indices $labelfirstlocation_x $($labelfirstlocation_y+$formgap*$indices))
    if($i -in (2))
    {
        New-Variable -Name $boxvariablelist[$i] -Value $(addobject "Textbox" $boxvariablelist[$i] $boxsize_x $boxsize_y $null $indices $boxfirstlocation_x $($boxfirstlocation_y+$formgap*$indices))
    }
    else 
    {
        New-Variable -Name $boxvariablelist[$i] -Value $(addobject "Combobox" $boxvariablelist[$i] $boxsize_x $boxsize_y $labletextlist[$i] $indices $boxfirstlocation_x $($boxfirstlocation_y+$formgap*$indices))
    }
}

#Region
[void] $lbregion.Items.AddRange($($importedcsv | Select-Object -Unique region).region)
$lbregion.Add_SelectedIndexChanged({
    readmsg "You have selected region as $($lbregion.text)." 1
    $lbcountry.Items.clear()
    $lbcountry.Items.AddRange($($importedcsv | Where-Object { $_.region -eq $($lbregion.text)}).country)
})

#Country
$lbcountry.Add_SelectedIndexChanged({
    readmsg "You have selected country as $($lbcountry.text)." 1
    findtag
})

#GroupTag
$tbgrouptag.enabled=$false

# Add ToolTips for the buttons
$tooltip1 = New-Object System.Windows.Forms.ToolTip
$ShowHelp={

     Switch ($this.name) {
        "okbutton"  {$tip = "Review the information and click submit to process the request."}
        "helpButton" {$tip = "Click to get help."}
        "CancelButton" {$tip = "Click to cancel the process and exit."}       
      }
     $tooltip1.SetToolTip($this,$tip)
   }

#okbutton
$okButton = addobject "Button" "okButton" 75 23 "Submit" $null 30 110
$form.AcceptButton = $okButton
$okButton.Add_MouseHover($ShowHelp)
$okButton.add_click({
    Start-Job -ScriptBlock {sethardwarehash $args[0] $args[1]} -InitializationScript $init -ArgumentList ($($tbgrouptag.text)),$tenentid
    #$form.close()
})

#cancelbutton
$cancelButton = addobject "Button" "cancelButton" 75 23 "Cancel" $null 130 110
$cancelButton.Add_MouseHover($ShowHelp)
$cancelButton.add_click({$form.close()})

#helpbutton
$helpButton = addobject "Button" "helpButton" 75 23 "Help" $null 230 110
$helpButton.Add_MouseHover($ShowHelp)
$helpButton.add_click({
$msgBoxInput=$null
$msgBoxInput = [System.Windows.MessageBox]::Show("", '!!! We are here to help you !!!','ok')
switch  ($msgBoxInput) 
{
    'OK' {}
}
})

#Error
$msg = New-Object System.Windows.Forms.RichTextBox
$msg.Location = New-Object System.Drawing.Point(10,15)
$msg.Size = New-Object System.Drawing.Size(380,125)
$msg.scrollbars='Both'
 
$groupBoxSingle.controls.AddRange(@($okButton,$helpButton,$cancelButton,$lregion,$lbregion,$lcountry,$lbcountry,$lgrouptag,$tbgrouptag,$label6))
$form.controls.AddRange(@($label8,$pictureBox,$groupBoxSingle,$groupBoxMsg))
$groupBoxMsg.Controls.AddRange(@($msg))
$form.Topmost = $true
$form.ShowDialog()
