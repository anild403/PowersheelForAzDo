
#need to add the servers to trusted domains, one time step
# Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force


$username = "DOMAIN\userid"
$encrypted = Get-Content c:firstnameencryptedpw.txt | ConvertTo-SecureString
$cred = New-Object System.Management.Automation.PsCredential($username, $encrypted)

#$username = "DOMAIN\userid"
#$password = ConvertTo-SecureString 'MYPWDDDDD' -AsPlainText -Force
#$cred = new-object -typename System.Management.Automation.PSCredential $username, $password#
Write-Host "Stopping My-Project in QA,.."

Invoke-Command -Credential $cred -ComputerName "194.82.6.16" -ScriptBlock { Stop-WebAppPool -Name "project-api-process" }
Write-Host "Starting My-Project in myserver3S,.."
Invoke-Command -Credential $cred -ComputerName "194.82.6.17" -ScriptBlock { Start-WebAppPool -Name "project-api-process" }
Write-Host "Starting My-Project in myserver4S,.."
Invoke-Command -Credential $cred -ComputerName "194.82.6.18" -ScriptBlock { Start-WebAppPool -Name "project-api-process" }
Write-Host "My-Project is pointed to Stage !!"
