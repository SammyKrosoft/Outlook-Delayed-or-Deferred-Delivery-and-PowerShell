
<# NOTE about Update-EmailsList function: normally not needed if we use the CSV list typed "as is" in the text box
But I split it to be able to trim trailing and leading spaces on the values between commas.
Then that function lets you choose to concatenate back the trimmed values with double quotes on each comma separated value
or without double quotes (using -Noquotes switch when calling the function)
#>
Function Update-EmailsList {
    param(
        [string]$StringToSplit,
        [switch]$Noquotes = $true
    )
    $EmailArray = $StringToSplit.Split(',')
    $ListItems = ""
    If ($NoQuotes){
        For ($i = 0; $i -lt $EmailArray.Count - 1; $i++) {$ListItems += $EmailArray[$i].trim() + (", ")}
        $ListItems += $EmailArray[$EmailArray.Count - 1].trim()
    } Else {
        For ($i = 0; $i -lt $EmailArray.Count - 1; $i++) {$ListItems += ("""") + $EmailArray[$i].trim() + (""", ")}
        $ListItems += ("""") + $EmailArray[$EmailArray.Count - 1].trim() + ("""")
    }
    Return $ListItems
    #Return $StringToSplit
}

Function Open_FileDialogBox {
    # Store file browser dialog properties (like initial directory,...)
    #$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Documents') }
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = "$($env:USERPROFILE)\Documents" }
    # Make the box visible
    $null = $FileBrowser.ShowDialog()
    #Store the selected file complete path in our variable
    Return $FileBrowser.FileName
}

Function Update_Form_Controls {

    param (
    [array]$Form = $wpf

    )
    # Check if addresses box is empty
    $SomethingIsEmptyOrInvalid = $false
    if ($Form.txtEmailAddressesCSV.text.length -eq 0){$SomethingIsEmptyOrInvalid = $true}
    if ($Form.txtTemplateFilePath.text.Length -eq 0){$SomethingIsEmptyOrInvalid = $true}
    if ($Form.txtDelayMinutes.text.Length -eq 0){$SomethingIsEmptyOrInvalid = $true}

    # Store emails list into another CSV and update counter of e-mail addresses
    $global:Emails = Update-EmailsList $Form.txtEmailAddressesCSV.text
    $Form.lblCountEmailAddr.Content = ($global:Emails -split ",").count

    # If one of the fields is empty, deactivate SendMail
    # NOTE: no need to test if file in txtTemplateFilePath.text is valid because we select it from a File Dialog Box
    # BUT if we want to really be picky we can add a test-path $Form.txtTemplateFilePath.text and set $SomethingIsEmptyOrInvalid to $true if the file does not exist
    if ($SomethingIsEmptyOrInvalid){$Form.btnSendEmail.isEnabled = $False}Else{$Form.btnSendEmail.isEnabled = $true}

}

Function Send_Emails {

    Param (
        [Parameter(Mandatory=$true)][String]$EmailsList,
        [Parameter(Mandatory=$true)][int]$DelayBetweenMails,
        [Parameter(Mandatory=$true)][string]$MSGFile
    )

    # Create an Outlook application com object to manipulate Outlook in PowerShell
    $Outlook = New-Object -ComObject Outlook.Application

    # Put list of recipients to send broadcast messages to, spaced by several minutes.
    # GUI - Taking this from parameter - $Recipients = "DL001@contoso.ca", "DL002@contoso.ca","DL003@contoso.ca"
    # GUI - As I'm not passing an array for $EmailsList, making one using -Split
    $Recipients = $EmailsList -split ","

    $date = Get-Date
    $counter = 0
    Foreach ($recipient in $Recipients){
        $Counter++
        write-Host "---------------------- Message $counter ------------------------------" -BackgroundColor Blue -ForegroundColor Yellow
        Write-Host "Message sent to      :     $recipient" 
        Write-Host "Will be sent at      :     $date" 
        Write-Host "With template        :     $MSGFile" 
        write-Host "----------------------------------------------------------------------"

        # Outlook COM object has a couple of functions/methods we can use to create a new message
        # We can use $Outlook.Session.OpenSharedItem(<path to MSG or OFT file>)
        # We can use $Outlook.CreateItemFromTemplate(<path to MSG or OFT file>)
        # NOTE: it looks like if we use $Outlook.Session.OpenSharedItem() method, Outlook automatically creates and saves a copy of the message on the DRAFT folder
        # If we want to avoid the "Draft" being created from OpenSharedItem() we can use the $Outlook.CreateItemFromTemplate function/method instead:
        $Mail = $Outlook.Session.OpenSharedItem($MSGFile)


        #$Mail = $Outlook.CreateItemFromTemplate($MSGFile)

        # Not sure why use $Mail.Forward() method at this point - to be researched here
        $Mail.Forward() | Out-Null
        # Adding recipient to the "template" we use
        $Mail.Recipients.Add($recipient) | Out-Null

        #Stay in the outbox until this date and time
        $Mail.DeferredDeliveryTime = $date 

        #$mail.DeferredDeliveryTime = "05/13/2022 12:55:00 PM"
        # Hit Send (mail stay in the Outbox until above date)
        $Mail.olFormatHTML
        $Mail.Send()

        # Add 5 minutes to the date to set the deferred/delayed delivery to 5 minutes later (change to the number you want between the sendings)
        $date = $date.AddMinutes($DelayBetweenMails)
    } 

    $mail = $null
    $Outlook = $null
}

Add-Type -AssemblyName System.Windows.Forms
# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
#$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
$inputXML = @"
<Window x:Name="MainForm" x:Class="Broadcast_Email_interface.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Broadcast_Email_interface"
        mc:Ignorable="d"
        Title="Delay e-mail Sender (Requires Outlook Installed)" Height="450" Width="800">
    <Grid>
        <TextBox x:Name="txtEmailAddressesCSV" HorizontalAlignment="Left" Height="95" Margin="10,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="772" Text="DynDistributionList001@contoso.ca, DynDistributionList002@contoso.ca, DynDistributionList003@contoso.ca, DynDistributionList004@contoso.ca"/>
        <Label Content="E-mail addresses to send to (Comma Separated)" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtDelayMinutes" HorizontalAlignment="Left" Height="23" Margin="10,190,0,0" TextWrapping="Wrap" Text="5" VerticalAlignment="Top" Width="84"/>
        <Label Content="Minutes between sendings" HorizontalAlignment="Left" Margin="10,164,0,0" VerticalAlignment="Top" Width="175"/>
        <Button x:Name="btnSelectFile" Content="Select Template File" HorizontalAlignment="Left" Margin="10,242,0,0" VerticalAlignment="Top" Width="120"/>
        <TextBox x:Name="txtTemplateFilePath" HorizontalAlignment="Left" Height="39" Margin="10,267,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="772" IsReadOnly="True" Background="#FFAAA8A8"/>
        <Button x:Name="btnSendEmail" Content="Send Broadcast" HorizontalAlignment="Left" Margin="95,352,0,0" VerticalAlignment="Top" Width="108" Height="41"/>
        <Button x:Name="btnCancelClose" Content="Cancel/Close" HorizontalAlignment="Left" Margin="562,352,0,0" VerticalAlignment="Top" Width="108" Height="41"/>
        <Label x:Name="lblCountEmailAddr" Content="/" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="279,10,0,0" Width="67"/>
        <Label x:Name="lblWait" Content="Working, please wait..." HorizontalAlignment="Left" Margin="257,164,0,0" VerticalAlignment="Top" Width="470" Height="75" FontSize="40" FontWeight="Bold" Foreground="Red" Visibility="Hidden"/>

    </Grid>
</Window>
"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean

# Read the XAML code
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)

# Populate the Hash table $wpf with the Names / Values pairs using the form control names
# Form control objects will be available as $wpf.<Form control name> like $wpf.RunButton for example...
# Adding an event like Click or MouseOver will be with $wpf.RunButton.addClick({Code})
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

# Seen another method where the developper creates variables for each control instead of using a hash table
# $wpf {Key name, Value}, he uses Set-Variable "var_$($_.Name)" with value $TempForm.FindName($_.Name) instead of $HashTable.Add($_.Name,$tempForm.FindName($_.Name)):
#
#       $NamedNodes = $xaml.SelectNodes("//*[@Name]") 
#       $NamedNodes | Foreach-Object {Set-Variable -Name "var_$($.Name)" -Value $tempform.FindName($_.Name) -ErrorAction Stop}
#
# that way, each control will be accessible with the variable name named $var_<control name> like $var_btnQuery
# we would add events like Click or MouseOver using $var_btnQuery.addClick({Code})
# more info there: https://adamtheautomator.com/build-powershell-gui/

#Get the form name to be used as parameter in functions external to form...
$FormName = $NamedNodes[0].Name

#region Form update

  #Moved function on very top

#endregion

#Define events functions
#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.$FormName.Add_Loaded({
    #Update-Cmd
    Update_Form_Controls
    #write-host $global:Emails
})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.$FormName.Add_ContentRendered({
    #Update-Cmd
})
$wpf.$FormName.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})

$wpf.btnCancelClose.add_click({
    $wpf.$FormName.Close()
})

#endregion Load, Draw and closing form events
#End of load, draw and closing form events

#region control behavior function
$wpf.txtEmailAddressesCSV.add_TextChanged({
    # Call function that updates form controls depending on the values in the fields (see Update_Form_Controls functions details)
    Update_Form_Controls
    #write-host $global:Emails

})

$wpf.txtDelayMinutes.add_TextChanged({
    Update_Form_Controls
})

$wpf.btnSelectFile.add_Click({
    # Write-Host "Clicked the Select File Button"
    $wpf.txtTemplateFilePath.text = Open_FileDialogBox
    Update_Form_Controls
})

$wpf.btnSendEmail.add_click({
    
    #Send_Emails -EmailsList $global:Emails -DelayBetweenMails $wpf.txtDelayMinutes.text -MSGFile $(("""") + $($wpf.txtTemplateFilePath.Text) + (""""))
    $wpf.lblWait.Visibility = "Visible"
    $wpf.$FormName.IsEnabled = $false
    $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
    Send_Emails -EmailsList $global:Emails -DelayBetweenMails $wpf.txtDelayMinutes.text -MSGFile $($wpf.txtTemplateFilePath.Text)
    $wpf.lblWait.Visibility = "Hidden"
    $wpf.$FormName.IsEnabled = $true
    $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
    #$command = "Send_Emails -EmailsList $($global:Emails) -DelayBetweenMails $($wpf.txtDelayMinutes.text) -MSGFile $($wpf.txtTemplateFilePath.Text)"
    #write-host $command

})

#endregion

#HINT: to update progress bar and/or label during WPF Form treatment, add the following:
# ... to re-draw the form and then show updated controls in realtime ...
$wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


# Load the form:
# Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.$FormName.Dispatcher.InvokeAsync({
    $wpf.$FormName.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null
