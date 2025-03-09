
#$groups = 'project_CCS_WC_Adjuster', 'project_CCS_WC_Supervisor', 'project_CCS_WC_Analyst','project_CCS_WC_Manager', 'project_CCS_WC_COPSAdmin'

$groups = 'project_CCS_Adjusters','project_CCS_Supervisors','project_CCS_Managers','project_CCS_WC_COPSAdmin'


foreach($user in Get-Content C:\test\data.txt) {

   foreach ($group in $groups) {
    $members = Get-ADGroupMember -Identity $group -Recursive | Select -ExpandProperty SamAccountName

    If ($members -contains $user) {
        Write-Host "$user is a member of $group"
    } Else {
        Write-Host "$user is not a member of $group"
    }
}

}