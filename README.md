# Outlook-PowerShell-send-mail-delay
Outlook PowerShell send mail delay

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

Another example, full script:
===============================

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

To change the default account:

```powershell
$Mail.SendUsingAccount = $ol.Session.Accounts | where {$_.DisplayName -eq $FromMail}
```

References:
[Create Outlook email draft using PowerShell](https://stackoverflow.com/questions/1453723/create-outlook-email-draft-using-powershell)

[Delay or schedule sending email messages](https://support.microsoft.com/en-us/office/delay-or-schedule-sending-email-messages-026af69f-c287-490a-a72f-6c65793744ba)

To use an OFT (Outlook Offline Template)
=========================================

```powershell
$outlook = New-Object -comObject Outlook.Application 
$mail = $outlook.Session.OpenSharedItem("C:\Temp\tmp.oft")
$mail.Forward()
$mail.Recipients.Add("DistributionList001@contoso.ca") 
$mail.send()
```
