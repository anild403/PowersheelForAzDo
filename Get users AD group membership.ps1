#Get-ADGroup -Filter {Name -like '*CCS*'}  -Properties * | select -property SamAccountName,Name,Description,DistinguishedName,CanonicalName,GroupCategory,GroupScope,whenCreated
 Get-ADPrincipalGroupMembership -Identity hooeyd1 |
      Out-String |
        Set-Content .\Process.txt