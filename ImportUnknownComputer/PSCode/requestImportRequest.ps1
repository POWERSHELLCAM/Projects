Function createCSV()
{
    $label6.ForeColor="black"
    $label6.Text = "Processing your request. Please wait......"
    Start-Sleep 3
    $macreg="^([0-9A-Fa-f]{2}[:]){5}([0-9A-Fa-f]{2})|([0-9a-fA-F]{4}\\.[0-9a-fA-F]{4}\\.[0-9a-fA-F]{4})$"
    $mailreg="^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
    $compreg="^(?![0-9]{1,15}$)[a-zA-Z0-9-_]{1,15}$"
    $guidreg='(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'
    $email=$country=$computer=$macaddress=$make=$model=$null
    $email="$($textBox0.text)"
    $country="$($listBox0.text)"
    $computer="$($textBox2.text)"
    $macaddress=$($textBox3.text)
    $guid=$($textBox3b.text)
    $make="$($listBox1.text)"
    $model="$($textBox5.text)"
    $label6.ForeColor="red"
    if(($email -ne  "") -and ($country -ne  "") -and ($computer -ne  "") -and (($macaddress -ne  "") -or ($guid -ne  "")) -and ($make -ne  "") -and ($model -ne  ""))
    {
        if(!(test-path "$sharedpath\csvdumps\$($computer).csv" -ErrorAction SilentlyContinue))
        {
            if(($macaddress -match $macreg) -or ($guid -match $guidreg))
            {
                if($email -match $mailreg)
                {
                    if ($Computer -match $compreg) 
                    {
                        try 
                        {
                            New-Item "$sharedpath\csvdumps\$($computer).csv" -ItemType File
                            $props=[ordered]@{
                                Email=$email
                                Country=$country
                                Computer=$computer
                                MACAddress=$macaddress
                                GUID=$guid
                                Make=$make
                                Model=$model
                        }
                            New-Object PsObject -Property $props | Export-Csv "$sharedpath\csvdumps\$($computer).csv" -NoTypeInformation
                            $msgBoxInput=$null
                            $msgBoxInput = [System.Windows.MessageBox]::Show("The device registration process will begin shortly. You will be notified on below email address about its progress. `n`n Email : $email", '!!! Good Luck !!!','ok')
                            switch  ($msgBoxInput) 
                            {
                                'OK' {$Form.Close() }
                            }
                        }
                        catch 
                        {
                            $label6.Text = "Error occured: $_ "
                        }
                        
                    }
                    else 
                    {
                        $label6.Text ="Computer name is not a valid one."
                    }
                }
                else 
                {
                    $label6.Text = "Email address is not in a valid format."
                }    
            }
            else 
            {
                $label6.Text = "Entered MACAddress or GUID is not in a valid format."
            }
        }
        else 
        {
            $label6.Text = "The computer name is already registered."
        }
       
    }
    else 
    {
        $label6.Text = "All fields are required one."
    }    
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles();
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$Forms = 'system.Windows.Forms.Form'
$Button = 'system.Windows.Forms.Button'
$FontStyle = 'Microsoft Sans Serif,10,style=Regular'
$SysDrawPoint = 'System.Drawing.Point'
$picture='Windows.Forms.PictureBox'
$sharedpath="\\<path>\ImportUnknownComputer\PSCode"
$file = (get-item "$sharedpath\_logo\.png")
$destPath = "$sharedpath\csvdumps"
$textboxsize="210"
$img = [System.Drawing.Image]::Fromfile($file);

#Main form design
$form = New-Object -TypeName $forms
$form.Text = '<company> MECM Unknown Device Registration'
$form.Size = New-Object System.Drawing.Size(750,400)
$form.StartPosition = 'CenterScreen'
$Form.BackColor="white"

$groupBoxSingle = New-Object System.Windows.Forms.GroupBox
$groupBoxSingle.Location = New-Object System.Drawing.Size(40,70) 
$groupBoxSingle.size = New-Object System.Drawing.Size(340,280) 
$groupBoxSingle.text = "Import Single Computer" 
$Form.Controls.Add($groupBoxSingle) 

$groupBoxMultiple = New-Object System.Windows.Forms.GroupBox
$groupBoxMultiple.Location = New-Object System.Drawing.Size(400,70) 
$groupBoxMultiple.size = New-Object System.Drawing.Size(320,180) 
$groupBoxMultiple.text = "Import Multiple Computer" 
$Form.Controls.Add($groupBoxMultiple ) 

# Add ToolTips for the buttons
$tooltip1 = New-Object System.Windows.Forms.ToolTip
$ShowHelp={

     Switch ($this.name) {
        "computername"  {$tip = "Exit Button INFO here."}
        "CopyButton" {$tip = "Copy Button INFO here."}
        "CancelButton" {$tip = "Clear Button INFO here."}       
      }
     $tooltip1.SetToolTip($this,$tip)
   }

$pictureBox = new-object -TypeName $picture
$pictureBox.Location = New-Object System.Drawing.Size(250,10)
$pictureBox.Size = New-Object System.Drawing.Size(500,50)
$pictureBox.Image = $img
$Form.controls.add($pictureBox)


$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(30,250)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$form.AcceptButton = $okButton


$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(130,250)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.name = 'cancelButton'
$form.CancelButton = $cancelButton
$cancelButton.Add_MouseHover($ShowHelp)


$helpButton = New-Object System.Windows.Forms.Button
$helpButton.Location = New-Object System.Drawing.Point(230,250)
$helpButton.Size = New-Object System.Drawing.Size(75,23)
$helpButton.Text = 'Help'
$helpButton.name = 'helpButton'

$helpButton.add_click({
$msgBoxInput=$null
$msgBoxInput = [System.Windows.MessageBox]::Show("1. Register device by MACAddress or BIOS GUID. Not by both. `n`n2. Computer name should be 15 character only.`n`n3. For multiple device registration, download the template first from `
from 'Import Multiple Computer' section. The template will be downloaded in user's download folder.`n`n4. Country field should be from the below list only. Country name must match exactly else it will fail. `n`tGlobal`n`tUS`n`tBelgium`n`tItaly`n`tUK`n`tJapan`
`n`n5. Upload the filled .csv file only. Other format not supported.`n`n6. The processing will take time based on `n`tCompletely new device : More time to process the request`n`tAlready registered : will be processed quickly`
`n`n7. You will receive the mail on mentioned mail id in .csv file about the processing.", '!!! We are here to help you !!!','ok')
switch  ($msgBoxInput) 
{
    'OK' {}
}
})

$cancelButton.add_click({$form.close()})
$okButton.add_click({$label6.Text = "";createCSV})

#Email Address
$label0 = New-Object System.Windows.Forms.Label
$label0.Location = New-Object System.Drawing.Point(10,20)
$label0.Size = New-Object System.Drawing.Size(100,20)
$label0.Text = 'Email Address : '
$textBox0 = New-Object System.Windows.Forms.TextBox
$textBox0.Location = New-Object System.Drawing.Point(120,20)
$textBox0.Size = New-Object System.Drawing.Size($textboxsize,20)

#Country
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,50)
$label1.Size = New-Object System.Drawing.Size(100,20)
$label1.Text = 'Country : '
$listBox0 = New-Object System.Windows.Forms.ComboBox
$listBox0.Location = New-Object System.Drawing.Point(120,50)
$listBox0.Size = New-Object System.Drawing.Size($textboxsize,20)
[void] $listBox0.Items.Add('Select country')
[void] $listBox0.Items.Add('Global')
[void] $listBox0.Items.Add('US')
[void] $listBox0.Items.Add('Belgium')
[void] $listBox0.Items.Add('Italy')
[void] $listBox0.Items.Add('UK')
[void] $listBox0.Items.Add('Japan')

#Computername
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,80)
$label2.Size = New-Object System.Drawing.Size(100,20)
$label2.Text = 'Computer Name : '
$label2.name = 'computername'
$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(120,80)
$textBox2.Size = New-Object System.Drawing.Size($textboxsize,20)


#MAc Address
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10,110)
$label3.Size = New-Object System.Drawing.Size(100,20)
$label3.Text = 'MAC Address : '
$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(120,110)
$textBox3.Size = New-Object System.Drawing.Size($textboxsize,20)
$label7 = New-Object System.Windows.Forms.Label
$label7.Location = New-Object System.Drawing.Point(120,130)
$label7.Size = New-Object System.Drawing.Size(210,20)
$label7.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",7,[System.Drawing.FontStyle]::Regular)
$label7.Text = '(Enter MACAddress value seperated by : only). Use either MAC or GUID option'

#GUID
$label3B = New-Object System.Windows.Forms.Label
$label3B.Location = New-Object System.Drawing.Point(10,160)
$label3B.Size = New-Object System.Drawing.Size(100,20)
$label3B.Text = 'BIOS GUID : '
$textBox3B = New-Object System.Windows.Forms.TextBox
$textBox3B.Location = New-Object System.Drawing.Point(120,160)
$textBox3B.Size = New-Object System.Drawing.Size($textboxsize,20)


# Make
$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(10,190)
$label4.Size = New-Object System.Drawing.Size(100,20)
$label4.Text = 'Make : '
$listBox1 = New-Object System.Windows.Forms.ComboBox
$listBox1.Location = New-Object System.Drawing.Point(120,190)
$listBox1.Size = New-Object System.Drawing.Size($textboxsize,20)
[void] $listBox1.Items.Add('Select make')
[void] $listBox1.Items.Add('HP')
[void] $listBox1.Items.Add('Dell')
[void] $listBox1.Items.Add('Lenovo')
[void] $listBox1.Items.Add('Acer')
[void] $listBox1.Items.Add('Asus')
[void] $listBox1.Items.Add('Microsoft')

#Model
$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(10,220)
$label5.Size = New-Object System.Drawing.Size(100,20)
$label5.Text = 'Model : '
$textBox5 = New-Object System.Windows.Forms.TextBox
$textBox5.Location = New-Object System.Drawing.Point(120,220)
$textBox5.Size = New-Object System.Drawing.Size($textboxsize,20)

#Error
$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(400,270)
$label6.Size = New-Object System.Drawing.Size(350,40)
$label6.Text = ""
$form.Controls.Add($label6)

$groupBoxSingle.controls.AddRange(@($label7,$label5,$label4,$label3,$label2,$label1,$label0,$okButton,$cancelButton,$label3B,$helpButton))
$groupBoxSingle.controls.AddRange(@($textBox0,$textBox1,$textBox2,$textBox3,$textBox3B,$textBox4,$textBox5,$textBox6,$listBox0,$listBox1))
$label8 = New-Object System.Windows.Forms.Label
$label8.Location = New-Object System.Drawing.Point(400,340)
$label8.Size = New-Object System.Drawing.Size(350,25)
$label8.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
$label8.Text = 'Developed by Shishir Kushawaha (Global Technology EUC Team)'
$form.Controls.Add($label8)


$txtFileName = New-Object -TypeName system.Windows.Forms.TextBox
$txtFileName.Size = New-Object System.Drawing.Size(300,20)
$txtFileName.location = New-Object -TypeName $SysDrawPoint -ArgumentList (10,20)
$txtFileName.Font = 'Microsoft Sans Serif,8,style=Regular'
 
$btnFileBrowser = New-Object -TypeName $Button
$btnFileBrowser.BackColor = '#1a80b6'
$btnFileBrowser.text = 'Browse'
$btnFileBrowser.Size = New-Object System.Drawing.Size(100,30)
$btnFileBrowser.location = New-Object -TypeName $SysDrawPoint -ArgumentList (20,60)
$btnFileBrowser.Font = $FontStyle
$btnFileBrowser.ForeColor = '#ffffff'
 
$btnUpload = New-Object -TypeName $Button
$btnUpload.BackColor = '#00FF7F'
$btnUpload.text = 'Upload'
$btnUpload.Size = New-Object System.Drawing.Size(100,30)
$btnUpload.location = New-Object -TypeName $SysDrawPoint -ArgumentList (140,60)
$btnUpload.Font = $FontStyle
$btnUpload.ForeColor = '#000000'

$lblFileName = New-Object -TypeName system.Windows.Forms.Label
$lblFileName.text = 'Use the below template to create a .csv file having mutliplte MACAddresses or BIOS GUIDs.'
$lblFileName.Size = New-Object System.Drawing.Size(300,30)
$lblFileName.location = New-Object -TypeName $SysDrawPoint -ArgumentList (10,100)


$LinkLabel = New-Object System.Windows.Forms.LinkLabel
$LinkLabel.Location = New-Object System.Drawing.Size(50,140)
$LinkLabel.Size = New-Object System.Drawing.Size(200,20)
$LinkLabel.LinkColor = "BLUE"
$LinkLabel.ActiveLinkColor = "RED"
$LinkLabel.Text = "Multiple MAC/GUID Template"
$LinkLabel.add_Click({$label6.Text = "Downloading template at $($HOME)\Downloads. Please wait.....";Start-BitsTransfer -Source "$sharedpath\_multipleMACSampleTemplate.csv" -Destination "$HOME\Downloads";$label6.ForeColor="red";$label6.Text = "File downloaded at '$($HOME)\Downloads'."})

 
#Adding the textbox,buttons to the forms for displaying
$groupBoxMultiple.controls.AddRange(@($txtFileName, $btnFileBrowser, $lblFileName, $btnUpload,$LinkLabel))
 
#Browse button click event
$btnFileBrowser.Add_Click({

#Creating an object for OpenFileDialog to a Form
$OpenDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
#Initiat browse path can be set by using initialDirectory
$OpenDialog.initialDirectory = $initialDirectory
#Set filter only to upload Excel file
$OpenDialog.filter = 'CSV Files (*.csv) |*.csv'
$OpenDialog.ShowDialog() | Out-Null
$filePath = $OpenDialog.filename
#Assigining the file choosen path to the text box
$txtFileName.Text = $filePath 
$TrackerFileUpload.Refresh()
})
#Upload button click eventy
$btnUpload.Add_Click({
#Set the destination path
$newFile="multipleMAC-$(Get-Date -Format 'yyyyMMdd-HHMMss').csv"
$destPath = "$sharedpath\csvdumps\$newFile"

try 
{
    Copy-Item -Path $txtFileName.Text -Destination $destPath
    $label6.Text = "File Uploaded successfully."
    $txtFileName.Text = ""
    $msgBoxInput=$null
    $msgBoxInput = [System.Windows.MessageBox]::Show("File uploaded successfully. The device registration process will begin shortly. You will be notified on email address mentioned in .csv file about its progress", '!!! Good Luck !!!','ok')
    switch  ($msgBoxInput) 
    {
        'OK' {}
    }
}
catch 
{
    write-host "$_"
    $label6.Text = "Failed to upload .CSV file."
}

})
 

$txtFileName.Text = ""

$form.Topmost = $true
$form.ShowDialog()