function displayRichText($text)
{
    $trich.appendtext("`n---->")
    $trich.appendtext($text)
    $trich.ScrollToend()
}

function displayMsg($m)
{
    $global:msg.text=""
    $global:msg.foreground='black'
    if(($m -match 'finished') -or ($m -match 'successfully'))
    {
        $global:msg.foreground='green'
    }
    elseif(($m -match 'started') -or ($m -match 'processing'))
    {
        $global:msg.foreground='darkblue'
    }
    elseif(($m -match 'not ') -or ($m -match 'no ') -or ($m -match 'error ') -or ($m -match 'failed'))
    {
        $global:msg.foreground='red'
    }
    else 
    {
        $global:msg.foreground='black'
    }
    $global:msg.text=$m
    displayRichText $m
}

[xml]$XAMLWindow = '
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Intunewin Application Utility" Height="480" Width="1000">
    <Window.Resources>
        <Style x:Key="RoundCorner" TargetType="{x:Type Button}">
            <Setter Property="HorizontalContentAlignment" Value="Center"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Grid x:Name="grid">
                            <Border x:Name="border" CornerRadius="8" BorderBrush="Black" BorderThickness="2">
                                <ContentPresenter HorizontalAlignment="Center"
                                          VerticalAlignment="Center"
                                          TextElement.FontWeight="Regular">
                                </ContentPresenter>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" TargetName="border">
                                    <Setter.Value>
                                        <RadialGradientBrush GradientOrigin="0.496,1.052">
                                            <RadialGradientBrush.RelativeTransform>
                                                <TransformGroup>
                                                    <ScaleTransform CenterX="0.5" CenterY="0.5" ScaleX="1.5" ScaleY="1.5"/>
                                                    <TranslateTransform X="0.02" Y="0.3"/>
                                                </TransformGroup>
                                            </RadialGradientBrush.RelativeTransform>
                                            <GradientStop Color="#00000000" Offset="1"/>
                                            <GradientStop Color="#FF303030" Offset="0.3"/>
                                        </RadialGradientBrush>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="BorderBrush" TargetName="border" Value="#FF33962B"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" TargetName="grid" Value="0.25"/>
                            </Trigger>

                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Height="550" VerticalAlignment="Top">
        <Label Content="Source Folder:" HorizontalAlignment="Left" Height="33" Margin="31,93,0,0" VerticalAlignment="Top" Width="105" FontWeight="Bold"/>
        <Label Content="Source File:" HorizontalAlignment="Left" Height="33" Margin="31,146,0,0" VerticalAlignment="Top" Width="105" FontWeight="Bold"/>
        <Label Content="Output Folder:" HorizontalAlignment="Left" Height="33" Margin="31,214,0,0" VerticalAlignment="Top" Width="106" RenderTransformOrigin="0.468,3.748" FontWeight="Bold"/>
        <Button Name="bcreate" Content="Create" HorizontalAlignment="Left" Height="29" Margin="154,304,0,0" VerticalAlignment="Top" Width="81" Style="{DynamicResource RoundCorner}"/>
        <Button Name="bcancel" Content="Cancel" HorizontalAlignment="Left" Height="29" Margin="401,304,0,0" VerticalAlignment="Top" Width="82" Style="{DynamicResource RoundCorner}" />
        <Button Name="bclear" Content="Clear" HorizontalAlignment="Left" Height="29" Margin="271,304,0,0" VerticalAlignment="Top" Width="88" Style="{DynamicResource RoundCorner}"/>
        <RadioButton Name="outputselection1" Content="Same as source Folder" HorizontalAlignment="Left" Height="33" Margin="156,220,0,0" VerticalAlignment="Top" Width="159" GroupName="outputfoler"/>
        <RadioButton Name="outputselection2" Content="Custom folder" HorizontalAlignment="Left" Height="33" Margin="336,220,0,0" VerticalAlignment="Top" Width="159" GroupName="outputfoler" IsChecked="True"/>
        <RichTextBox Name="rich" HorizontalAlignment="Left" Height="300" Margin="590,93,0,0" VerticalAlignment="Top" Width="365" />
        <TextBox Name="tsourcefolder" HorizontalAlignment="Left" Height="33" Margin="156,93,0,0" TextWrapping="Wrap" Text="Browse source folder of application." VerticalAlignment="Top" Width="318"/>
        <TextBox Name="tsourcefile" HorizontalAlignment="Left" Height="33" Margin="156,143,0,0" TextWrapping="Wrap" Text="Browse setup file from source folder" VerticalAlignment="Top" Width="318"/>
        <TextBox Name="toutput" HorizontalAlignment="Left" Height="33" Margin="156,256,0,0" TextWrapping="Wrap" Text="Browse output folder" VerticalAlignment="Top" Width="318"/>
        <Button Name="bsourcefolder" Content="Browse" HorizontalAlignment="Left" Height="33" Margin="484,93,0,0" VerticalAlignment="Top" Width="83"  />
        <Button Name="bsourcefile" Content="Browse" HorizontalAlignment="Left" Height="33" Margin="484,143,0,0" VerticalAlignment="Top" Width="83"  />
        <Button Name="boutput" Content="Browse" HorizontalAlignment="Left" Height="33" Margin="484,256,0,0" VerticalAlignment="Top" Width="83"  />
        <Separator HorizontalAlignment="Left" Height="10" Margin="6,70,0,0" VerticalAlignment="Top" Width="970"/>
        <Label Content="Package Name:" HorizontalAlignment="Left" Height="34" Margin="31,359,0,0" VerticalAlignment="Top" Width="110" FontWeight="Bold"/>
        <Label Name="pname" Content="" HorizontalAlignment="Left" Height="34" Margin="135,359,0,0" VerticalAlignment="Top" Width="150"/>
        <Label Content="Package Size:" HorizontalAlignment="Left" Height="34" Margin="313,359,0,0" VerticalAlignment="Top" Width="103" FontWeight="Bold"/>
        <Label Name="psize" Content="Package size output" HorizontalAlignment="Left" Height="34" Margin="405,359,0,0" VerticalAlignment="Top" Width="150"/>
        <StatusBar HorizontalAlignment="Left" Height="35" Margin="1,408,0,0" VerticalAlignment="Top" Width="980" >
            <StatusBar.ItemsPanel>
                <ItemsPanelTemplate>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="132" />
                            <ColumnDefinition Width="Auto" />
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="Auto" />
                            <ColumnDefinition Width="100" />
                            <ColumnDefinition Width="Auto" />
                            <ColumnDefinition Width="210" />
                        </Grid.ColumnDefinitions>
                    </Grid>
                </ItemsPanelTemplate>
            </StatusBar.ItemsPanel>
            <StatusBarItem>
                <TextBlock Name="clock" Text="06-05-2023 13:25" Width="190" />
            </StatusBarItem>
            <Separator Grid.Column="1" />
            <StatusBarItem Grid.Column="2">
                <TextBlock Name="msg" Text="Welcome to Intunewin Win32 Application creation utility." />
            </StatusBarItem>
            <Separator Grid.Column="3" />
            <StatusBarItem Grid.Column="4">
                <ProgressBar Name="progress" Value="5" Width="90" Height="16" HorizontalAlignment="Center" />
            </StatusBarItem>
            <Separator Grid.Column="5" />
            <StatusBarItem Grid.Column="6">
                <TextBlock Text="Developed by Shishir Kushawaha" />
            </StatusBarItem>
        </StatusBar>
        <Label Content="Easily convert EXE/MSI to IntuneWin32 Application" HorizontalAlignment="Center" Height="60" Margin="0,7,0,0" VerticalAlignment="Top" Width="980" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="36" FontWeight="Bold" Background="#FFF97B42" Foreground="White"/>
    </Grid>
</Window>
'
clear-host
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles();
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")


# Create the Window Object
$Reader=(New-Object System.Xml.XmlNodeReader $XAMLWindow)
$Window=[Windows.Markup.XamlReader]::Load( $Reader )
$scriptpath=$PSScriptRoot
$Progress = $Window.FindName('progress')
$clock = $Window.FindName('clock')
$bsourcefile=$Window.FindName("bsourcefile")
$bsourcefolder=$Window.FindName("bsourcefolder")
$boutputfolder=$Window.FindName("boutput")
$bcreate=$Window.FindName("bcreate")
$bclear=$Window.FindName("bclear")
$bcancel=$Window.FindName("bcancel")
$trich=$Window.FindName("rich")
$tsourcefile=$Window.FindName("tsourcefile")
$tsourcefile.isenabled=$false
$tsourcefolder=$Window.FindName("tsourcefolder")
$toutputfolder=$Window.FindName("toutput")
$pname=$Window.FindName("pname")
$psize=$Window.FindName("psize")
$choiceSourceFolder=$Window.FindName("outputselection1")
$choiceCustomFolder=$Window.FindName("outputselection2")
$global:msg = $Window.FindName('msg')
$global:sourceFolderPath=$global:sourceFilePath=$global:outputFolderPath=$Null

$sb={
    $stdout=$stderr=$ExitCode=$null     
    if(Test-Path "$($args[0])\IntuneWinAppUtil.exe" -ErrorAction SilentlyContinue)
    {
        if($args[1] -and $args[2] -and $args[3])
        {
            $arguments = "-c `"$($args[3])`" -s `"$($args[2])`" -o `"$($args[1])`" -q"
            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = "$($args[0])\IntuneWinAppUtil.exe"
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.UseShellExecute = $false
            $pinfo.Arguments = $arguments
            $p = New-Object System.Diagnostics.Process
            $p.StartInfo = $pinfo
            $p.Start() | Out-Null
            $p.WaitForExit()
            $stderr=$null
            $stdout = $p.StandardOutput.ReadToEnd()
            $stderr = $p.StandardError.ReadToEnd()
            $ExitCode=$p.ExitCode
        }
        else
        {
            $stderr+="Can't create intunewin application as either of the parameter is not defined."
        }
    }
    else
    {
        $stderr+="Can't create intunewin application as IntuneWinAppUtil.exe does not exists in root folder."
    }
    return "`nProcessOutput: $stdout","`nProcessError: $stderr","`nProcessExitCode: $ExitCode"
}

#Browse button click event
$bsourcefolder.Add_Click({
    $tsourcefile.isenabled=$false
    $tsourcefile.text=$global:sourceFilePath=$Null
    displayMsg "Processing package source folder."
    $openSourceFolderDiaglog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
    $openSourceFolderDiaglog.Description = "Select package source folder."
    $openSourceFolderDiaglog.rootfolder = "MyComputer"
    $openSourceFolderDiaglog.SelectedPath = $initialDirectory
    if($openSourceFolderDiaglog.ShowDialog() -eq "OK")
    {
        $global:sourceFolderPath = $openSourceFolderDiaglog.SelectedPath
    }
    $tsourcefolder.text= $sourceFolderPath
    $Progress.Value=10
    displayMsg "Source folder: $sourceFolderPath"
    $psize.content="{0:N2} MB" -f ((Get-ChildItem $sourceFolderPath -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB) 
    if($choiceSourceFolder.IsChecked)
    {
        $global:outputFolderPath =$sourceFolderPath
        $toutputfolder.text= $outputFolderPath 
    }
})

$bsourcefile.Add_Click({
    displayMsg "Processing package source file."
    $tsourcefile.isenabled=$true
    $openSourceFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $openSourceFileDialog.initialDirectory = $sourceFolderPath
    $openSourceFileDialog.filter = 'Executable Files|*.exe;*.msi'
    $openSourceFileDialog.ShowDialog() | Out-Null
    $global:sourceFilePath = $openSourceFileDialog.filename
    $tsourcefile.text = $sourceFilePath
    if($sourceFolderPath -ne ([System.IO.DirectoryInfo]"$sourceFilePath").parent.fullname)
    {
        displayMsg "The package file is not in the package source folder."
    }
    else
    {
        $Progress.Value=30
        displayMsg "Source file: $sourceFilePath"
        $pname.Content=$(([System.IO.DirectoryInfo]"$sourceFilePath").name)
    }
})

$boutputfolder.Add_Click({
    displayMsg "Processing output folder."
    $openOutputFolderDiaglog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
    $openOutputFolderDiaglog.Description = "Select output folder."
    $openOutputFolderDiaglog.rootfolder = "MyComputer"
    $openOutputFolderDiaglog.SelectedPath = $initialDirectory
    if($openOutputFolderDiaglog.ShowDialog() -eq "OK")
    {
        $global:outputFolderPath = $openOutputFolderDiaglog.SelectedPath
    }
    $toutputfolder.text= $outputFolderPath 
    $Progress.Value=50
    displayMsg "Output folder: $outputFolderPath "
})

$choiceSourceFolder.add_Checked({
    if($null -ne $sourceFolderPath)
    {
        $boutputfolder.isenabled=$toutputfolder.isenabled= $false
        $toutputfolder.text=$sourceFolderPath
        $Progress.Value=50
        $global:outputFolderPath =$sourceFolderPath
        $toutputfolder.text= $outputFolderPath 
        displayMsg "Output folder: $outputFolderPath "
    }
    else
    {
        displayMsg "No package source folder is defined. Please set source folder first."
    }
})

$bcreate.Add_Click({
    displayMsg "Started ituneWin32 application creation. Please wait...."
    $bcreate.content="Please wait..."
    Start-Job -name packagingJob -ScriptBlock $sb -ArgumentList $scriptpath,$outputFolderPath,$sourceFilePath,$sourceFolderPath

    displayMsg "Waiting for packaging job to finish."
    wait-Job -name packagingJob
    $packagingJobResult=receive-job -name packagingJob
    $bcreate.content="Create"
    if($packagingJobResult[0] -match 'has been generated successfully')
    {
        $Progress.Value=100
        displayRichText $packagingJobResult[0]
        displayMsg "Successfully created $(([System.IO.DirectoryInfo]"$sourceFilePath").name) intunewin application."
    }
    else
    {
            displayMsg "Failed to create $(([System.IO.DirectoryInfo]"$sourceFilePath").name) intunewin application."
            displayRichText $packagingJobResult[1]
            displayRichText $packagingJobResult[2]
    }             
})

$bcancel.Add_Click({
    $Window.close()
})

$bclear.Add_Click({
    $global:sourceFolderPath=$global:sourceFilePath=$global:outputFolderPath=$Null
    $tsourcefile.text=$tsourcefolder.text=$toutputfolder.text=""
})

$choiceCustomFolder.add_Checked({
    $boutputfolder.isenabled=$True
    $toutputfolder.text=""
})

$timer1 = New-Object 'System.Windows.Forms.Timer'
$timer1_Tick={
    $clock.text= (Get-Date).ToString("dd-MM-yyyy HH:mm:ss")
}
$timer1.Enabled = $True
$timer1.Interval = 1
$timer1.add_Tick($timer1_Tick)

# Open the Window
$Window.ShowDialog() | Out-Null
