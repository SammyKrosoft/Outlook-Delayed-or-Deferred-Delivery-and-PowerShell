# Outlook Delayed Delivery

## Intro

This repository gives information and links about Oultook delayed delivery (sending messages but asking Outlook to process these at a later date/time). 

Below are a couple articles links and extracts about the behavior difference between sending messages with delayed delivery on Outlook configured with cache mode, and on Outlook configured without cache mode (aka online mode).

Then I threw in a few links and PowerShell code samples about how to automate Outlook with Powershell to send a delayed message from a MSG or OFT template. First I show a few code samples to explain different parts of Outlook automation with Powershell (delayed delivery, creating new e-mail from a MSG or OFT template, etc..).

Finally, this repository contains a complete PowerShell script with a Graphical User Interface using WPF where users can:

- put the e-mail addresses to send messages to
- put the delay they want to set between Outlook messages sendings (delayed deliveries)
- choose the template they want to use to send these messages

Some details and screenshots about this WPF GUI Outlook automation script [can be found on the README_Outlook_BroadcastInterface.md](https://github.com/SammyKrosoft/Outlook-PowerShell-send-mail-delay/blob/main/README_Outlook_BroadCastInterface.md)

## Details

[From Stilstick article](https://www.slipstick.com/outlook/delay-sending-message-outlook-closed/)

```
When you use an Microsoft Exchange mailbox (either on-prem or Office 365 Exchange online) you can send messages later, with Outlook closed, as long as you can set the account up in Outlook desktop using Online mode.

When you use online mode, the deferred messages are submitted to the Exchange message queue and held until the scheduled time as Outlook doesn’t have a local cache when you use online mode, so it can’t hold them at the client.

This only works if you have the account set up in Outlook in online mode. It will not work with Outlook.com accounts as they do not support online mode.

By default, Outlook sets Exchange accounts up in cached mode. You can check in File, Account Settings: open the account settings dialog, and double click on the account. Cached will be ticked by default when you set up the account, untick it to drop to online mode. The status bar should say 'Online', when you have cached mode turned off.

Before you disable cached mode: if you have the messages already in the Outbox move them to Drafts folder first so they will sync to the server. After you switch to online mode, go to the Drafts folder and Send the messages.

You won't see the messages if you look in Outlook on the web, so you just have to trust Exchange. To verify it is working, create an email to email address you can check on your phone. Defer it for 15 min from now and click Send. Close Outlook. If you receive the message in 15 min, it’s set up properly.
```

[From this Howto-Outlook.com article](https://www.howto-outlook.com/howto/schedule-recurring-email.htm)

```
The PowerShell code in this guide allows you to send an email template, that you’ve created in Outlook.

The Send-OutlookMailFromTemplate PowerShell script allows you to send a message template created in Outlook (oft-file). This will give you the full message composing and formatting capabilities of Outlook and the ability to easily edit it if needed.

Optionally, the script allows you to add an attachment at the time of sending to make sure that the file you want to add is always up-to-date without the need to modify the Outlook template.

The script can easily be scheduled in Windows Task Scheduler and can be configured with a recurrence pattern.
```

## Using PowerShell to automate Outlook

Here's a nice, complete and simple to understand process to user PowerShell with Outlook:

https://community.spiceworks.com/how_to/150253-send-mail-from-powershell-using-outlook


## Some code samples

[From StackOverflow](https://stackoverflow.com/questions/14809023/sending-defer-message-delivery-and-change-default-account-using-powershell)

With the below you'll get a list of all methods/properties available on the mail object:

```powershell
$ol = New-Object -comObject Outlook.Application  
$mail = $ol.CreateItem(0)  
$mail | Get-Member
```

One property is DeferredDeliveryTime. 
You can find info on this link as well [MailItem.DeferredDeliveryTime property (Outlook)](https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.deferreddeliverytime)

You can set it like this:

```powershell
#Stay in the outbox until this date and time
$mail.DeferredDeliveryTime = "6/6/2022 10:50:00 AM"
```

Or:

```powershell
#Wait 10 minutes before sending mail
$date = Get-Date
$date = $date.AddMinutes(10)
$mail.DeferredDeliveryTime = $date
```

## Another example, full script to send a message through Outlook automation with PowerShell with a delivery 10 minutes from the time user sent the message 

```powershell
$ol = New-Object -comObject Outlook.Application 
$ns = $ol.GetNameSpace("MAPI")

# call the save method yo dave the email in the drafts folder
$mail = $ol.CreateItem(0)
$null = $Mail.Recipients.Add("DistributionList001@contoso.ca")  
$Mail.Subject = "PS1 Script TestMail"  
$Mail.Body = "  Test Mail  "

$date = Get-Date
$date = $date.AddMinutes(10)
$Mail.DeferredDeliveryTime = $date #"06/06/2022 10:50:00 AM"

$Mail.save()

# get it back from drafts and update the body
$drafts = $ns.GetDefaultFolder($olFolderDrafts)
$draft = $drafts.Items | where {$_.subject -eq 'PS1 Script TestMail'}
$draft.body += "`n adding text"
$draft.save()

$inspector = $draft.GetInspector  
$inspector.Display()


# send the message
$draft.Send()
```

If you need to change the default account to use with Outlook:

```powershell
$Mail.SendUsingAccount = $ol.Session.Accounts | where {$_.DisplayName -eq $FromMail}
```

**NOTE:** The above line is to use a different account, not a different profile. If you want to be able to open Outlook with a specific profiles, you need to get the list of profiles from the local registry (or if you know the name of your profile you can use it as well hard coded in your script), and then open your Outlook session using ``` $Outlook.Session.Logon("Profile Name") ``` 

Here's a script to list all the accounts under each profile by [Jeremy Corbello](https://social.technet.microsoft.com/profile/jacorbello/?ws=usercard-mini) [from an answer to the Microsoft communities](https://social.technet.microsoft.com/Forums/ie/en-US/c9b80a2c-9775-4299-a348-1faa16757b68/retrieve-outlook-info-account-from-multi-profile?forum=winserverpowershell) :

```powershell
$profiles = (Get-ChildItem HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles).PSChildName

foreach ($profile in $profiles) {
    $Outlook = New-Object -ComObject Outlook.Application
    $Outlook.Session.Logon("$profile")
    $NameSpace = $ol.GetNamespace("MAPI")
    $NameSpace.Accounts | ft Displayname,Username,SMTPAddress -AutoSize
    } 

```

So before sending a mail, and if you don't have Outlook already open, just after creating your Outlook COM object in PowerShell, open the session using a profile name either manually as shown below, or from the registry as shown above:

```powershell
$Outlook = New-Object -ComObject Outlook.Application
$Outlook.Session.Logon("Outlook Profile 01")

$mail = $Outlook.CreateItem(0)

# or to open from a template (MSG or OFT) - the below seems to add the Outlook signature:

$mail = $Outlook.CreateItemFromTemplate("path to the template - can be an OFT or a MSG")

# or to open from a MSG e-mail, similar to opening from a template like above - doesn't seem to add the Outlook signature:

$mail = $Outlook.Session.OpenSharedItem("path to the MSG file")

```



And finally if you want to Quit Outlook and cleanup your variable:

- Quit Outlook (no need if you want to keep Outlook opened):

```powershell
$Outlook.Quit()
```

- Cleanup your variable from the COM object to free up memory

It's not enough to set the $Outlook variable to $null, the below 3 code lines is a common practice used to remove a variable containing a COM object:

```powershell
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
```

### References:

[Create Outlook email draft using PowerShell](https://stackoverflow.com/questions/1453723/create-outlook-email-draft-using-powershell)

[Delay or schedule sending email messages](https://support.microsoft.com/en-us/office/delay-or-schedule-sending-email-messages-026af69f-c287-490a-a72f-6c65793744ba)

## To use an OFT (Outlook Offline Template)

```powershell
 $PathToOft = "c:\temp\Bulkmessage001.oft"
 $mailboxes = "DistributionList001@contoso.ca"
    
 $outlook = New-Object -comObject Outlook.Application 
 $mail = $Outlook.CreateItemFromTemplate("$PathToOft")
    
 foreach ($mailbox in $mailboxes){
        
     $mail.Forward()
     $mail.Recipients.Add($mailbox)
     $mail.To = $mailbox
     $mail.Save()
     $mail.send()
 }
```

## Putting all the above together: Script sample to send an e-mail from a .MSG or .OFT template to several recipients (DLs for example), adding 5 minutes between each sending.

```powershell
cls
# Previously "hard coded" the path to the .msg or .oft file in a variable. Instead used Windows.Forms.OpenFileDialog to select the file.
# $MSGFile = "$($env:UserProfile)\Desktop\AutomatedMessage.msg"

# The below routine is to use Windows open file dialog to select a file to use as the broadcast message.
#Loading Windows Form "assembly" (=library)
Add-Type -AssemblyName System.Windows.Forms
# Store file browser dialog properties (like initial directory,...)
#$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Documents') }
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = "$($env:USERPROFILE)\Documents" }
# Make the box visible
$null = $FileBrowser.ShowDialog()
#Store the selected file complete path in our variable
$MSGFile = $FileBrowser.FileName


# Create an Outlook application com object to manipulate Outlook in PowerShell
$Outlook = New-Object -ComObject Outlook.Application

# Put list of recipients to send broadcast messages to, spaced by several minutes.
$Recipients = "DL001@contoso.ca", "DL002@contoso.ca","DL003@contoso.ca"

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
    #$Mail = $Outlook.Session.OpenSharedItem($MSGFile)
    $mail = $Outlook.CreateItemFromTemplate("$MSGFile")

    # Not sure why use $Mail.Forward() method at this point - to be researched here
    $Mail.Forward() | Out-Null
    # Adding recipient to the "template" we use
    $Mail.Recipients.Add($recipient) | Out-Null

    #Stay in the outbox until this date and time
    $Mail.DeferredDeliveryTime = $date 

    #$mail.DeferredDeliveryTime = "05/13/2022 12:55:00 PM"
    # Hit Send (mail stay in the Outbox until above date)
    $Mail.Send() 

    # Add 5 minutes to the date to set the deferred/delayed delivery to 5 minutes later (change to the number you want between the sendings)
    $date = $date.AddMinutes(5)
} 


$mail = $null
$Outlook = $null
```
