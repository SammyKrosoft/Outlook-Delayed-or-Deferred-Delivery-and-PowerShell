# Graphical Interface to Send Messages to multiple users (broadcast e-mails)

**Download link at the end of this article (Right-click Save As to download the latest .ps1 version of this script)**

Here's a GUI to help out sending multiple e-mails from a single template - for more information about how to use Outlook COM object with PowerShell to send e-mails, see the Readme.md on this repository.

![image](https://user-images.githubusercontent.com/33433229/173739480-fc9b5d8c-303b-4392-836e-83fd48bd0cd6.png)

NOTE: the “Send Broadcast” button is greyed out if at least one of the fields is empty (e-mail addresses, minutes between sendings or template file). 
NOTE2: note that when you add or remove e-mail addresses, the number increments or decrements dynamically:

![image](https://user-images.githubusercontent.com/33433229/173739595-3214b62f-2060-43d5-9477-8c8fdfb99eee.png)

Note the number changing if we remove 2 of the addresses (and commas):

![image](https://user-images.githubusercontent.com/33433229/173739613-080a512f-0abb-4115-a874-7f5235991791.png)

-	Choose the minutes between sendings if you don’t like the default:

![image](https://user-images.githubusercontent.com/33433229/173739665-d350f375-630f-42c1-9304-9e9a4a827f22.png)

-	Click on “Select Template File” to open the file selection dialog box (just reused the Open File Dialog we used in the non-GUI version of the script), select the file and click “Open”:

![image](https://user-images.githubusercontent.com/33433229/173739780-a6236cbf-c83d-4faf-999c-6134165237d2.png)

![image](https://user-images.githubusercontent.com/33433229/173739785-4925b8f6-c091-4faf-9c5e-ce53563f3f64.png)

![image](https://user-images.githubusercontent.com/33433229/173739797-61c70488-ce4c-460a-a71c-b3175251d5aa.png)

![image](https://user-images.githubusercontent.com/33433229/173739808-13ef6e5e-e420-4188-8d92-76dfbdaa4d2c.png)

![image](https://user-images.githubusercontent.com/33433229/173739816-358a0ec6-2db1-457d-a73c-55f1eb3e3d0d.png)

- When the script is working in the background, the interface will deactivate with the "Working, please wait..." message:

![image](https://user-images.githubusercontent.com/33433229/173739914-ad81349c-711c-46a1-9642-6f8edf6170ec.png)

- Also, while the script runs to send the deferred delivery messages, you'll see some information regarding each message sent on the underlying Powershell window:

![image](https://user-images.githubusercontent.com/33433229/173851840-98817ea3-d243-4355-b83b-0b71e400c67c.png)

- After execution, the interface returns to an active state:

![image](https://user-images.githubusercontent.com/33433229/173739956-0577c866-eb98-4e34-add8-3b5c1aa38cb3.png)

-	And check in Outlook’s Outbox – the first message is sent immediately, and the others will be sent on the interval set on the form 

NOTE: the “Sent” date you see on the Outbox will be the current date and time, but the delivery is deferred by the number of minutes you set on the form

NOTE2: if you open one of these messages, it won’t be processed anymore – you will have to click “Send” to put them back in the Outlook Outbox processing queue.

![image](https://user-images.githubusercontent.com/33433229/173739981-3bf34a29-5ea8-4df9-9584-60b3e2d564fe.png)

-	Finally to close the form, click “Cancel/Close” or the “X”on the top right corner

![image](https://user-images.githubusercontent.com/33433229/173740001-81d848e6-7cd5-4fae-a763-7f752a873722.png)

![image](https://user-images.githubusercontent.com/33433229/173740007-cf35a337-0f05-4d49-bf50-a25de07a1a62.png)



# Download (Right-Click - Save link as)

**Right-click *Save link as*** ==> [Download Send Bulk E-mail PowerShell Script](https://raw.githubusercontent.com/SammyKrosoft/Outlook-PowerShell-send-mail-delay/main/Outlook_BroadCastInterface.ps1) <== **Right-click *Save link As***
