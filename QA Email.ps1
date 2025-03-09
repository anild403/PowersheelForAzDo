
$FromAddress = "firstname middlename lastname<firstnamemiddlename.lastname@yourcompany.com>"
$Subject="My-Project in test will be pointed to QA <EOM>"
$MessageBody = @"
Hi All,

To support the testing, we are planning to point the My-Project QA environment to Oracle TCLM database and turn it off the in STAGING environment.

This will propagate the data changes in MCS Test to QA queues instead of the STAGING queues and may impact the functionalities dependent on My-Project in STAGING environment.

Please let us know if there are any concerns, I will let you know when the change is completed.

Thanks,
firstname middlename lastname.
ITS Team Leader
ITS Application Delivery - project 
yourcompany Insurance Company of America
973 948 1788

"@
#create COM object named Outlook
$Outlook = New-Object -ComObject Outlook.Application
#create Outlook MailItem named Mail using CreateItem() method
$Mail = $Outlook.CreateItem(0)
#add properties as desired
$Mail.To = $FromAddress
$Mail.Subject = $Subject
$Mail.Body = $MessageBody
#send message
$Mail.Send()
#quit and cleanup
#$Outlook.Quit()
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

Write-Host "Email sent for QA !!"