$username = "DOMAIN\userid"
$encrypted = Get-Content C:\powershell\firstnameencryptedpw.txt | ConvertTo-SecureString
$cred = New-Object System.Management.Automation.PsCredential($username, $encrypted)

do{



$a = Invoke-Command -Credential $cred -ComputerName '194.82.2.199' -ScriptBlock {Get-WebAppPoolState}

ForEach ($item in $a){
    if($item.value -eq 'Stopped')
    {
        
        Write-Host $item.itemxpath
        #create COM object named Outlook
        $Outlook = New-Object -ComObject Outlook.Application
        #create Outlook MailItem named Mail using CreateItem() method
        $Mail = $Outlook.CreateItem(0)
        #add properties as desired
        $Mail.To = "firstnamemiddlename.lastname@yourcompany.com"
        $Mail.CC = "Robert.Labar@yourcompany.com"
        $Mail.BCC = "Rajesh.Kannappan@yourcompany.com"
        $Mail.Subject = "App pool " + ($item.itemxpath -split ("'"))[1] + " is down in QA"
        $Mail.Body = ($item.itemxpath -split ("'"))[1] + " is down in QA"
        #send message
        $Mail.Send()
        Write-Host "Sent email!"

    }
}
    start-sleep -Seconds 900

}until($infinity)

